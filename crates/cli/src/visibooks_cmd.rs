//! `vgrid visibooks`: the API key a VisiBooks recipe source reads with, and
//! the entities that key can reach.

use std::io::{BufRead, IsTerminal, Write};

use visigrid_io::recipe::visibooks::{self, DEFAULT_SERVER};

use crate::CliError;

/// Save (or with `delete`, remove) the key for `server` in the keychain. The
/// key is read from stdin, unechoed at a terminal, so it never appears in
/// shell history or the process list.
pub(crate) fn cmd_key(server: Option<String>, delete: bool) -> Result<(), CliError> {
    let origin = visibooks::origin(server.as_deref().unwrap_or(DEFAULT_SERVER)).map_err(CliError::args)?;
    let account = visibooks::keychain_account(&origin);
    if delete {
        visigrid_config::secrets::delete(&account).map_err(CliError::io)?;
        eprintln!("Removed the VisiBooks key for {origin}");
        return Ok(());
    }
    let key = read_secret(&format!("VisiBooks API key for {origin} (from Settings > API keys): "))?;
    if key.is_empty() {
        return Err(CliError::args("no key given; nothing saved"));
    }
    visigrid_config::secrets::set(&account, &key).map_err(CliError::io)?;
    eprintln!("Saved in the system keychain. Recipes reading {origin} will use it.");
    Ok(())
}

/// List the entities the saved key reaches, with the ids recipes name.
pub(crate) fn cmd_entities(server: Option<String>) -> Result<(), CliError> {
    let origin = visibooks::origin(server.as_deref().unwrap_or(DEFAULT_SERVER)).map_err(CliError::args)?;
    let entities = visibooks::entities(&origin).map_err(CliError::io)?;
    if entities.is_empty() {
        eprintln!("The key reaches no entities; grant it some in VisiBooks (Settings > API keys).");
    }
    for (id, name) in entities {
        println!("{id}\t{name}");
    }
    Ok(())
}

fn read_secret(prompt: &str) -> Result<String, CliError> {
    let stdin = std::io::stdin();
    if !stdin.is_terminal() {
        let mut line = String::new();
        stdin.lock().read_line(&mut line).map_err(|e| CliError::io(e.to_string()))?;
        return Ok(line.trim().to_string());
    }
    eprint!("{prompt}");
    let _ = std::io::stderr().flush();
    use crossterm::event::{read, Event, KeyCode, KeyEventKind, KeyModifiers};
    crossterm::terminal::enable_raw_mode().map_err(|e| CliError::io(e.to_string()))?;
    let mut key = String::new();
    let result = loop {
        match read() {
            Ok(Event::Key(k)) if k.kind != KeyEventKind::Release => match k.code {
                KeyCode::Enter => break Ok(()),
                KeyCode::Char('c') if k.modifiers.contains(KeyModifiers::CONTROL) => break Err(CliError::args("cancelled")),
                KeyCode::Backspace => {
                    key.pop();
                }
                KeyCode::Char(c) => key.push(c),
                _ => {}
            },
            Ok(Event::Paste(text)) => key.push_str(&text),
            Ok(_) => {}
            Err(e) => break Err(CliError::io(e.to_string())),
        }
    };
    let _ = crossterm::terminal::disable_raw_mode();
    eprintln!();
    result.map(|_| key.trim().to_string())
}
