//! `vgrid serve --share`: push frames to the live-session relay so a browser
//! (or a phone) can watch this workbook.
//!
//! The host pushes over plain HTTPS — `vgrid serve` listens on localhost,
//! which no browser and no server can reach — and VisiAPI fans frames out to
//! viewers. Outward only: viewers have no write path in phase 1.
//!
//! Sharing is an ADDITION to serving, never a dependency of it. Every network
//! failure here is logged and swallowed; the local session keeps working.

use std::time::{Duration, Instant};

use serde_json::json;
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::SheetLayout;

/// Send a heartbeat if nothing else has gone out in this long. The relay
/// treats a session as stale after 2 minutes.
const HEARTBEAT: Duration = Duration::from_secs(60);

/// The relay rejects frames larger than this; a snapshot that would exceed it
/// is skipped with a warning rather than failing the session.
const MAX_FRAME_BYTES: usize = 2 * 1024 * 1024;

pub struct ShareSession {
    client: reqwest::blocking::Client,
    api_base: String,
    token: String,
    code: String,
    pub url: String,
    last_send: Instant,
}

impl ShareSession {
    /// Open a live session. Returns None (with an explanation on stderr) if
    /// sharing can't start — serving continues either way.
    pub fn open(title: &str) -> Option<Self> {
        let creds = match visigrid_hub_client::load_auth() {
            Some(c) => c,
            None => {
                eprintln!("--share needs an API key: run `vgrid login` first. Serving locally only.");
                return None;
            }
        };
        let client = reqwest::blocking::Client::builder()
            .timeout(Duration::from_secs(10))
            // reqwest sends NO User-Agent by default, and the API's edge
            // answers a UA-less request with a bare 403 — no body, nothing
            // Rails ever sees. Identify ourselves.
            .user_agent(concat!("vgrid/", env!("CARGO_PKG_VERSION")))
            .build()
            .ok()?;

        let resp = client
            .post(format!("{}/api/live_sessions", creds.api_base))
            .bearer_auth(&creds.token)
            .json(&json!({ "title": title }))
            .send();

        let body: serde_json::Value = match resp {
            Ok(r) if r.status().is_success() => r.json().ok()?,
            Ok(r) => {
                // Surface the server's own explanation — it is far more
                // actionable than the status code ("Bearer vk_* API key
                // required" vs "403 Forbidden").
                let status = r.status();
                let detail = r
                    .json::<serde_json::Value>()
                    .ok()
                    .and_then(|b| {
                        b.get("error")
                            .or_else(|| b.get("message"))
                            .and_then(|v| v.as_str().map(str::to_string))
                    })
                    .unwrap_or_else(|| status.to_string());
                eprintln!("--share could not open a live session: {}", detail);
                if status.as_u16() == 401 || status.as_u16() == 403 {
                    eprintln!("  Live sessions need a vk_ API key: create one in Settings, then `vgrid login`.");
                }
                eprintln!("  Serving locally only.");
                return None;
            }
            Err(e) => {
                eprintln!("--share could not reach {} ({}). Serving locally only.", creds.api_base, e);
                return None;
            }
        };

        let code = body.get("code")?.as_str()?.to_string();
        let url = body
            .get("url")
            .and_then(|u| u.as_str())
            .map(str::to_string)
            .unwrap_or_else(|| format!("https://app.visigrid.app/live/{}", code));

        Some(Self {
            client,
            api_base: creds.api_base,
            token: creds.token,
            code,
            url,
            last_send: Instant::now(),
        })
    }

    fn post_frame(&mut self, frame: serde_json::Value) {
        let body = json!({ "frame": frame });
        let encoded = match serde_json::to_vec(&body) {
            Ok(v) => v,
            Err(e) => {
                eprintln!("share: could not encode frame ({})", e);
                return;
            }
        };
        if encoded.len() > MAX_FRAME_BYTES {
            eprintln!(
                "share: frame is {} KB, over the {} KB relay limit — skipped (viewer may lag until the next change)",
                encoded.len() / 1024,
                MAX_FRAME_BYTES / 1024
            );
            return;
        }
        let res = self
            .client
            .post(format!("{}/api/live_sessions/{}/frames", self.api_base, self.code))
            .bearer_auth(&self.token)
            .header("content-type", "application/json")
            .body(encoded)
            .send();
        match res {
            Ok(r) if r.status().is_success() => {
                self.last_send = Instant::now();
            }
            // Never fatal: the local session is the product; sharing is extra.
            Ok(r) => {
                let status = r.status();
                let detail = r
                    .json::<serde_json::Value>()
                    .ok()
                    .and_then(|b| b.get("error").and_then(|v| v.as_str().map(str::to_string)))
                    .unwrap_or_else(|| status.to_string());
                eprintln!("share: relay rejected a frame: {}", detail);
            }
            Err(e) => eprintln!("share: frame not delivered ({})", e),
        }
    }

    /// Full state. Sent on open, and after any change a delta can't express
    /// (structural edits, sheets added or renamed).
    ///
    /// Layouts are passed through, not defaulted: the viewer renders column
    /// widths, frozen panes, and charts from them, and an empty layout ships
    /// a workbook missing half of what the host has on screen.
    pub fn send_snapshot(&mut self, wb: &Workbook, layouts: &[SheetLayout]) {
        let workbook = match visigrid_io::json::export_workbook(wb, layouts, wb.active_sheet_index()) {
            Ok(s) => s,
            Err(e) => {
                eprintln!("share: could not serialize workbook ({})", e);
                return;
            }
        };
        let workbook: serde_json::Value = match serde_json::from_str(&workbook) {
            Ok(v) => v,
            Err(e) => {
                eprintln!("share: workbook is not valid JSON ({})", e);
                return;
            }
        };
        self.post_frame(json!({
            "kind": "snapshot",
            "revision": wb.revision(),
            "workbook": workbook,
        }));
    }

    /// Changed cells. `raw` is what the engine holds (formula text or
    /// literal); `display` is the formatted value the viewer shows.
    pub fn send_delta(&mut self, wb: &Workbook, cells: &[(usize, usize, usize)]) {
        if cells.is_empty() {
            return;
        }
        let payload: Vec<serde_json::Value> = cells
            .iter()
            .filter_map(|(sheet, row, col)| {
                let s = wb.sheets().get(*sheet)?;
                Some(json!({
                    "sheet": sheet,
                    "row": row,
                    "col": col,
                    "raw": s.get_raw(*row, *col),
                    "display": s.get_display(*row, *col),
                }))
            })
            .collect();
        self.post_frame(json!({
            "kind": "delta",
            "revision": wb.revision(),
            "cells": payload,
        }));
    }

    /// Keep the session off the stale list while nothing is happening.
    pub fn heartbeat_if_idle(&mut self, wb: &Workbook) {
        if self.last_send.elapsed() >= HEARTBEAT {
            self.post_frame(json!({
                "kind": "delta",
                "revision": wb.revision(),
                "cells": [],
            }));
        }
    }

    /// Tell viewers the host ended. Their last state stays on screen.
    pub fn close(&mut self) {
        let _ = self
            .client
            .delete(format!("{}/api/live_sessions/{}", self.api_base, self.code))
            .bearer_auth(&self.token)
            .send();
    }
}
