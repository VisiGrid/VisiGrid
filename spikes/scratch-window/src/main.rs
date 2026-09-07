//! Phase 0 spike for VisiGrid scratch mode.
//!
//! Proves, on the pinned gpui fork, that a frameless PopUp window can be
//! summoned by a global hotkey while another app is frontmost, takes keyboard
//! focus, receives typed text, closes on Esc, and reopens with its text intact.
//! It also reports which display it landed on and how long summon took.
//!
//! Drive it three ways: the hotkey (Ctrl+Shift+Space), `kill -USR1 <pid>` on
//! unix, or typing `t` + Enter on stdin. `q` + Enter quits. `d` + Enter cycles
//! the target display. Every event is logged to stdout with a timestamp.

use gpui::{prelude::*, *};
use gpui_platform::application;
use global_hotkey::{
    hotkey::{Code, HotKey, Modifiers},
    GlobalHotKeyEvent, GlobalHotKeyManager, HotKeyState,
};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use std::sync::Arc;
use std::time::{Duration, Instant};

static TOGGLE_REQUESTED: AtomicBool = AtomicBool::new(false);
static QUIT_REQUESTED: AtomicBool = AtomicBool::new(false);
static CYCLE_DISPLAY: AtomicBool = AtomicBool::new(false);
static SUMMONS: AtomicUsize = AtomicUsize::new(0);

fn ts() -> String {
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default();
    format!("{:>6}.{:03}", now.as_secs() % 100000, now.subsec_millis())
}

fn log(msg: impl AsRef<str>) {
    println!("[{}] {}", ts(), msg.as_ref());
}

/// State that must survive the window being closed and reopened.
struct ScratchState {
    text: String,
    window: Option<WindowHandle<Scratch>>,
    summon_started: Option<Instant>,
    display_index: usize,
    last_active: Option<bool>,
}
impl Global for ScratchState {}

struct Scratch {
    focus_handle: FocusHandle,
    painted: bool,
}

impl Scratch {
    fn on_key(&mut self, ev: &KeyDownEvent, window: &mut Window, cx: &mut Context<Self>) {
        let ks = &ev.keystroke;
        match ks.key.as_str() {
            "escape" => {
                log("key: escape -> hide (remove_window)");
                window.remove_window();
                cx.global_mut::<ScratchState>().window = None;
                return;
            }
            "backspace" => {
                cx.global_mut::<ScratchState>().text.pop();
            }
            "enter" => {
                cx.global_mut::<ScratchState>().text.push('⏎');
            }
            _ => {
                if let Some(ch) = ks.key_char.as_ref() {
                    cx.global_mut::<ScratchState>().text.push_str(ch);
                } else if ks.modifiers.control && ks.key == "q" {
                    log("key: ctrl-q -> quit");
                    cx.quit();
                    return;
                }
            }
        }
        log(format!(
            "key: {:?} char={:?} active={}",
            ks.key,
            ks.key_char,
            window.is_window_active()
        ));
        cx.notify();
    }
}

impl Render for Scratch {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let active = window.is_window_active();
        let display = window.display(cx).map(|d| format!("{:?}", d.id()));
        let scale = window.scale_factor();
        let bounds = window.bounds();
        {
            let st = cx.global_mut::<ScratchState>();
            if !self.painted {
                self.painted = true;
                let elapsed = st.summon_started.map(|t| t.elapsed().as_millis()).unwrap_or(0);
                log(format!(
                    "first render {}ms after summon; display={:?} scale={} bounds={:?} active={}",
                    elapsed, display, scale, bounds, active
                ));
            }
            if st.last_active != Some(active) {
                st.last_active = Some(active);
                log(format!("window active={}", active));
            }
        }
        let text = cx.global::<ScratchState>().text.clone();
        let summons = SUMMONS.load(Ordering::Relaxed);
        div()
            .id("scratch-root")
            .track_focus(&self.focus_handle)
            .on_key_down(cx.listener(Self::on_key))
            .size_full()
            .flex()
            .flex_col()
            .gap_2()
            .p_4()
            .bg(rgb(0x061206))
            .text_color(rgb(0x4af626))
            .font_family("monospace")
            .text_sm()
            .child(format!(
                "scratch spike · summon #{summons} · active={active} · display={display:?} · scale={scale}"
            ))
            .child(
                div()
                    .text_xl()
                    .text_color(rgb(0x7cfc00))
                    .child(format!("> {text}_")),
            )
            .child("type · Esc hides · Ctrl+Q quits · hotkey Ctrl+Shift+Space reopens")
    }
}

fn target_display(cx: &App) -> Option<std::rc::Rc<dyn PlatformDisplay>> {
    let displays = cx.displays();
    if displays.is_empty() {
        return cx.primary_display();
    }
    let idx = cx.global::<ScratchState>().display_index % displays.len();
    displays.into_iter().nth(idx)
}

fn toggle(cx: &mut App) {
    let open = cx.global::<ScratchState>().window;
    if let Some(handle) = open {
        let still_open = cx.windows().iter().any(|w| w.window_id() == handle.window_id());
        if still_open {
            log("toggle: window open -> closing");
            let _ = handle.update(cx, |_, window, _| window.remove_window());
            cx.global_mut::<ScratchState>().window = None;
            return;
        }
    }
    let n = SUMMONS.fetch_add(1, Ordering::Relaxed) + 1;
    let started = Instant::now();
    cx.global_mut::<ScratchState>().summon_started = Some(started);

    let display = target_display(cx);
    let (origin, width) = match &display {
        Some(d) => {
            let vb = d.visible_bounds();
            let width = px((f32::from(vb.size.width) * 0.8).min(1200.0));
            let x = vb.origin.x + (vb.size.width - width) / 2.0;
            (point(x, vb.origin.y), width)
        }
        None => (point(px(100.0), px(0.0)), px(900.0)),
    };
    let bounds = Bounds { origin, size: size(width, px(360.0)) };
    log(format!(
        "toggle: summon #{n} kind={} on display {:?} bounds={:?}",
        std::env::var("SPIKE_KIND").unwrap_or_else(|_| "popup".into()),
        display.as_ref().map(|d| d.id()),
        bounds
    ));

    let options = WindowOptions {
        window_bounds: Some(WindowBounds::Windowed(bounds)),
        titlebar: None,
        focus: true,
        show: true,
        kind: match std::env::var("SPIKE_KIND").as_deref() {
            Ok("floating") => WindowKind::Floating,
            Ok("normal") => WindowKind::Normal,
            _ => WindowKind::PopUp,
        },
        is_movable: false,
        is_resizable: false,
        is_minimizable: false,
        display_id: display.as_ref().map(|d| d.id()),
        window_decorations: Some(WindowDecorations::Client),
        app_id: Some("visigrid-scratch-spike".into()),
        ..Default::default()
    };
    match cx.open_window(options, |window, cx| {
        let view = cx.new(|cx| Scratch { focus_handle: cx.focus_handle(), painted: false });
        let fh = view.read(cx).focus_handle.clone();
        window.focus(&fh, cx);
        view
    }) {
        Ok(handle) => {
            cx.global_mut::<ScratchState>().window = Some(handle);
            // Bring the app forward even if another app was frontmost (macOS).
            cx.activate(true);
            let _ = handle.update(cx, |view, window, cx| {
                window.activate_window();
                let fh = view.focus_handle.clone();
                window.focus(&fh, cx);
                cx.notify();
            });
            log(format!("toggle: open_window returned in {}ms", started.elapsed().as_millis()));
        }
        Err(e) => log(format!("toggle: open_window FAILED: {e:?}")),
    }
}

#[cfg(unix)]
extern "C" fn on_sigusr1(_: libc::c_int) {
    TOGGLE_REQUESTED.store(true, Ordering::SeqCst);
}

fn main() {
    log(format!("pid {}", std::process::id()));
    #[cfg(unix)]
    unsafe {
        libc::signal(libc::SIGUSR1, on_sigusr1 as libc::sighandler_t);
    }

    // stdin driver: t = toggle, q = quit, d = cycle display
    std::thread::spawn(|| {
        let stdin = std::io::stdin();
        let mut line = String::new();
        loop {
            line.clear();
            if stdin.read_line(&mut line).unwrap_or(0) == 0 {
                break;
            }
            match line.trim() {
                "t" => TOGGLE_REQUESTED.store(true, Ordering::SeqCst),
                "q" => QUIT_REQUESTED.store(true, Ordering::SeqCst),
                "d" => CYCLE_DISPLAY.store(true, Ordering::SeqCst),
                _ => {}
            }
        }
    });

    // Global hotkey. Registration failure is logged and everything else still works.
    let manager = match GlobalHotKeyManager::new() {
        Ok(m) => Some(m),
        Err(e) => {
            log(format!("hotkey: manager unavailable: {e} (use SIGUSR1 or stdin 't')"));
            None
        }
    };
    let hotkey = HotKey::new(Some(Modifiers::CONTROL | Modifiers::SHIFT), Code::Space);
    if let Some(m) = &manager {
        match m.register(hotkey) {
            Ok(()) => log("hotkey: registered ctrl+shift+space"),
            Err(e) => log(format!("hotkey: REGISTRATION FAILED: {e}")),
        }
    }
    let manager = Arc::new(manager); // keep alive for the app lifetime

    application().run(move |cx: &mut App| {
        let _keep = manager.clone();
        cx.set_quit_mode(QuitMode::Explicit);
        cx.set_global(ScratchState {
            text: String::new(),
            window: None,
            summon_started: None,
            display_index: 0,
            last_active: None,
        });
        log(format!(
            "app running; displays={}",
            cx.displays().iter().map(|d| format!("{:?}:{:?}", d.id(), d.bounds())).collect::<Vec<_>>().join(" ")
        ));

        cx.spawn(async move |cx: &mut AsyncApp| {
            let receiver = GlobalHotKeyEvent::receiver();
            loop {
                cx.background_executor().timer(Duration::from_millis(25)).await;
                while let Ok(ev) = receiver.try_recv() {
                    if ev.state() == HotKeyState::Pressed {
                        log("hotkey: pressed");
                        TOGGLE_REQUESTED.store(true, Ordering::SeqCst);
                    }
                }
                if CYCLE_DISPLAY.swap(false, Ordering::SeqCst) {
                    let _ = cx.update(|cx| {
                        let st = cx.global_mut::<ScratchState>();
                        st.display_index += 1;
                        log(format!("display index -> {}", st.display_index));
                    });
                }
                if TOGGLE_REQUESTED.swap(false, Ordering::SeqCst) {
                    let _ = cx.update(|cx| toggle(cx));
                }
                if QUIT_REQUESTED.swap(false, Ordering::SeqCst) {
                    let _ = cx.update(|cx| {
                        log("quit");
                        cx.quit();
                    });
                }
            }
        })
        .detach();
    });
}
