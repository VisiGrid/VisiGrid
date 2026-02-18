//! PTY lifecycle: spawn, event proxy, shutdown.

use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::Arc;
use std::thread::JoinHandle;

use alacritty_terminal::event::{Event, EventListener, WindowSize};
use alacritty_terminal::event_loop::{EventLoop, EventLoopSender};
use alacritty_terminal::sync::FairMutex;
use alacritty_terminal::term::{Config as TermConfig, Term};
use alacritty_terminal::tty::{self, Options as PtyOptions};

/// Event proxy that forwards terminal events to the UI via an mpsc channel.
#[derive(Clone)]
pub struct TerminalEventProxy {
    sender: std::sync::mpsc::Sender<Event>,
}

impl TerminalEventProxy {
    pub fn new(sender: std::sync::mpsc::Sender<Event>) -> Self {
        Self { sender }
    }
}

impl EventListener for TerminalEventProxy {
    fn send_event(&self, event: Event) {
        let _ = self.sender.send(event);
    }
}

/// Helper struct implementing `Dimensions` for initial Term creation.
struct TermDimensions {
    cols: usize,
    lines: usize,
}

impl alacritty_terminal::grid::Dimensions for TermDimensions {
    fn total_lines(&self) -> usize {
        self.lines
    }

    fn screen_lines(&self) -> usize {
        self.lines
    }

    fn columns(&self) -> usize {
        self.cols
    }
}

/// Spawn a new PTY session.
///
/// Returns the terminal instance, event loop sender, and the I/O thread handle.
pub fn spawn_pty(
    cwd: Option<PathBuf>,
    cols: u16,
    rows: u16,
    proxy: TerminalEventProxy,
) -> std::io::Result<(Arc<FairMutex<Term<TerminalEventProxy>>>, EventLoopSender, JoinHandle<()>)> {
    // Determine shell
    let shell = std::env::var("SHELL").unwrap_or_else(|_| "/bin/bash".to_string());

    // Environment variables
    let mut env = HashMap::new();
    env.insert("TERM".to_string(), "xterm-256color".to_string());
    env.insert("COLORTERM".to_string(), "truecolor".to_string());

    // Clear CLAUDECODE env var so AI CLIs (claude) don't refuse to start
    // when VisiGrid itself was launched from within a Claude Code session.
    std::env::remove_var("CLAUDECODE");

    let pty_options = PtyOptions {
        shell: Some(tty::Shell::new(shell, Vec::new())),
        working_directory: cwd,
        env,
        ..Default::default()
    };

    let window_size = WindowSize {
        num_lines: rows,
        num_cols: cols,
        cell_width: 8,   // Approximate; PTY only uses num_lines/num_cols for TIOCSWINSZ
        cell_height: 16,
    };

    // Create terminal config
    let config = TermConfig::default();

    // Create terminal
    let dimensions = TermDimensions {
        cols: cols as usize,
        lines: rows as usize,
    };
    let term = Term::new(config, &dimensions, proxy.clone());
    let term = Arc::new(FairMutex::new(term));

    // Create PTY
    let pty = tty::new(&pty_options, window_size, 0)?;

    // Create event loop
    let event_loop = EventLoop::new(
        Arc::clone(&term),
        proxy,
        pty,
        false, // drain_on_exit
        false, // ref_test
    )?;

    let sender = event_loop.channel();

    // Spawn I/O thread (returns JoinHandle<(EventLoop, State)>)
    let io_join = event_loop.spawn();

    // Wrap in a simple JoinHandle<()> — we don't need the return values
    let handle = std::thread::Builder::new()
        .name("terminal-pty-join".to_string())
        .spawn(move || {
            let _ = io_join.join();
        })
        .expect("failed to spawn terminal join thread");

    Ok((term, sender, handle))
}
