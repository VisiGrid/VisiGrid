//! PostgreSQL recipe sources against a real server: the read-only proof
//! accepts a SELECT-only role and refuses every way a role could write.
//!
//! Runs when VISIGRID_TEST_POSTGRES holds a superuser connection to a
//! server on this computer without TLS (CI: a postgres service container),
//! like `host=127.0.0.1 port=5432 user=postgres password=postgres dbname=postgres`.
//! It creates its own database and roles and drops them after.
//! One test, because the password is passed in process-wide PG* variables.

use std::path::Path;

use visigrid_io::recipe::{self, postgres as pg, Recipe};

struct Scratch {
    admin: String,
    db: String,
    roles: Vec<String>,
}

impl Drop for Scratch {
    fn drop(&mut self) {
        if let Ok(mut c) = postgres::Client::connect(&self.admin, postgres::NoTls) {
            let _ = c.batch_execute(&format!("DROP DATABASE IF EXISTS {} WITH (FORCE)", self.db));
            for r in &self.roles {
                let _ = c.batch_execute(&format!("DROP ROLE IF EXISTS {r}"));
            }
        }
    }
}

fn recipe(host: &str, port: u16, db: &str, user: &str, what: &str) -> Recipe {
    Recipe::from_toml(&format!(
        "version = 1\n[source]\nkind = \"postgres\"\nhost = \"{host}\"\nport = {port}\ndatabase = \"{db}\"\nuser = \"{user}\"\ntls = false\n{what}\n"
    ))
    .unwrap()
}

fn read(r: &Recipe, user: &str) -> Result<recipe::RunResult, String> {
    std::env::set_var("PGUSER", user);
    let snap = r.read_snapshot(Path::new("."), None)?;
    Ok(recipe::run(r, &snap))
}

#[test]
fn the_read_only_proof_accepts_a_reader_and_refuses_every_writer() {
    let Ok(admin) = std::env::var("VISIGRID_TEST_POSTGRES") else {
        eprintln!("VISIGRID_TEST_POSTGRES not set; skipping the PostgreSQL source test");
        return;
    };
    let cfg: postgres::Config = admin.parse().unwrap();
    let host = match &cfg.get_hosts()[0] {
        postgres::config::Host::Tcp(h) => h.clone(),
        #[allow(unreachable_patterns)]
        _ => panic!("VISIGRID_TEST_POSTGRES must name a TCP host"),
    };
    let port = cfg.get_ports().first().copied().unwrap_or(5432);
    let tag = format!("vgt{}", std::process::id());
    let role = |r: &str| format!("{tag}_{r}");
    let scratch = Scratch {
        admin: admin.clone(),
        db: tag.clone(),
        roles: ["reader", "writer", "member", "super", "creator"].iter().map(|r| role(r)).collect(),
    };

    let mut a = postgres::Client::connect(&admin, postgres::NoTls).unwrap();
    a.batch_execute(&format!("CREATE DATABASE {}", scratch.db)).unwrap();
    let pw = "test-password-not-secret";
    for r in &scratch.roles {
        a.batch_execute(&format!("CREATE ROLE {r} LOGIN PASSWORD '{pw}' NOINHERIT")).unwrap();
    }
    let mut d = postgres::Client::connect(&format!("{admin} dbname={}", scratch.db), postgres::NoTls).unwrap();
    // As the plan provisions it (Supabase runs the same SQL)
    d.batch_execute(&format!(
        "REVOKE CREATE ON SCHEMA public FROM PUBLIC;
         REVOKE CREATE, TEMP ON DATABASE {db} FROM PUBLIC;
         GRANT CONNECT ON DATABASE {db} TO PUBLIC;
         CREATE TABLE orders (id bigint PRIMARY KEY, number text, created_at timestamptz, total numeric(12,2),
                              big_ref bigint, gift boolean, tags text[], ship_date date,
                              tally bigint GENERATED ALWAYS AS (id * 2) STORED);
         INSERT INTO orders VALUES
           (1, 'LL-1001', '2026-09-01 12:00+02', 129.50, 9007199254740993, true, '{{a,b}}', '2026-09-03'),
           (2, 'LL-1002', '2026-09-02 11:30+00', 42.00, 12, false, '{{}}', NULL);
         CREATE VIEW paid AS SELECT id, number FROM orders WHERE total > 50;
         CREATE TABLE secret_notes (id int PRIMARY KEY, body text);
         INSERT INTO secret_notes VALUES (1, 'hidden');
         ALTER TABLE secret_notes ENABLE ROW LEVEL SECURITY;
         ALTER ROLE {reader} SET default_transaction_read_only = on;
         GRANT USAGE ON SCHEMA public TO {reader}, {writer}, {member}, {creator};
         GRANT SELECT ON ALL TABLES IN SCHEMA public TO {reader}, {writer}, {member}, {creator};
         GRANT UPDATE ON orders TO {writer};
         ALTER ROLE {writer} SET default_transaction_read_only = on;
         GRANT {writer} TO {member};
         ALTER ROLE {super} SUPERUSER;
         GRANT CREATE ON SCHEMA public TO {creator};",
        db = scratch.db,
        reader = role("reader"),
        writer = role("writer"),
        member = role("member"),
        super = role("super"),
        creator = role("creator"),
    ))
    .unwrap();

    std::env::set_var("PGHOST", &host);
    std::env::set_var("PGPORT", port.to_string());
    std::env::set_var("PGPASSWORD", pw);
    let table = |t: &str| format!("table = \"{t}\"");

    // A SELECT-only role reads, with each type kept
    let reader = role("reader");
    let r = recipe(&host, port, &scratch.db, &reader, &table("public.orders"));
    let out = read(&r, &reader).unwrap();
    assert!(out.report.ok, "{:?}", out.report.failures);
    let names: Vec<_> = out.output.columns.iter().map(|c| c.name.as_str()).collect();
    assert_eq!(names, ["id", "number", "created_at", "total", "big_ref", "gift", "tags", "ship_date", "tally"]);
    assert_eq!(out.output.rows[0][2], "2026-09-01 10:00:00", "timestamptz read in UTC");
    assert_eq!(out.output.rows[0][3], "129.50");
    assert_eq!(out.output.rows[0][4], "9007199254740993", "int8 past 2^53 keeps every digit");
    assert_eq!(out.output.rows[0][5], "TRUE");
    assert_eq!(out.output.rows[0][6], "{a,b}");
    assert_eq!(out.output.rows[1][7], "", "NULL date is empty");

    // The write was really tried, on the source table, past its generated column
    let mut src = match &r.source {
        recipe::Source::Postgres(s) => s.clone(),
        _ => unreachable!(),
    };
    std::env::set_var("PGUSER", &reader);
    let v = pg::verify(&src).unwrap();
    assert_eq!(v.role, reader);
    assert!(v.write_tried_on.is_some(), "{v:?}");

    // A view and a query read too. A view can't take the write test, so it
    // is tried on a table the role reads instead
    let view = recipe(&host, port, &scratch.db, &reader, &table("paid"));
    let out = read(&view, &reader).unwrap();
    assert_eq!(out.output.rows, [["1", "LL-1001"]]);
    let recipe::Source::Postgres(view_src) = &view.source else { unreachable!() };
    let (_, verified) = pg::fetch(view_src, 1000, 100).unwrap();
    assert!(verified.write_tried_on.is_some(), "{verified:?}");
    let q = "query = \"\"\"\nselect number, total * 2 as twice from orders -- comment\norder by id\n\"\"\"";
    let out = read(&recipe(&host, port, &scratch.db, &reader, q), &reader).unwrap();
    assert_eq!(out.output.rows[1], ["LL-1002", "84.00"]);

    // A query is one SELECT, and never changes data
    let two = "query = \"select 1; select 2\"";
    let e = read(&recipe(&host, port, &scratch.db, &reader, two), &reader).err().unwrap();
    assert!(e.contains("single SELECT"), "{e}");
    let delete = "query = \"delete from orders returning *\"";
    assert!(read(&recipe(&host, port, &scratch.db, &reader, delete), &reader).is_err());
    let cte = "query = \"with d as (delete from orders returning *) select * from d\"";
    assert!(read(&recipe(&host, port, &scratch.db, &reader, cte), &reader).is_err());

    // Row-level security hiding every row says so
    let out = read(&recipe(&host, port, &scratch.db, &reader, &table("secret_notes")), &reader).unwrap();
    assert_eq!(out.output.rows.len(), 0);
    assert!(out.report.warnings.iter().any(|w| w.contains("row-level security") && w.contains("CREATE POLICY")), "{:?}", out.report.warnings);

    // Every role that could write is refused, with what let it
    let refused = |who: &str, expect: &str| {
        let user = role(who);
        let e = read(&recipe(&host, port, &scratch.db, &user, &table("public.orders")), &user).err().unwrap();
        assert!(e.contains("isn't read-only") && e.contains(expect), "{who}: {e}");
    };
    refused("writer", "table orders: UPDATE");
    refused("member", &format!("member of {}", role("writer")));
    refused("super", "superuser");
    refused("creator", "schema public: CREATE");

    // A wrong password is named as such
    std::env::set_var("PGUSER", &reader);
    src.port = port;
    std::env::set_var("PGPASSWORD", "");
    let e = pg::verify(&src).err().unwrap();
    assert!(e.contains("no password saved") || e.contains("password"), "{e}");
    drop(scratch);
}
