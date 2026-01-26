//! License dialog for entering and viewing license status.

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, DialogFrame, DialogSize};

/// User-friendly license summary for UI display
/// Decoupled from raw license payload to avoid leaking internals
#[derive(Debug, Clone)]
pub struct LicenseDisplayInfo {
    pub edition: String,
    pub status: LicenseStatus,
    pub expires: String,
    pub key_id: String,
    pub plan: String,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LicenseStatus {
    Active,
    GracePeriod,
    Expired,
    Invalid,
    Free,
}

impl LicenseStatus {
    pub fn label(&self) -> &'static str {
        match self {
            LicenseStatus::Active => "Active",
            LicenseStatus::GracePeriod => "Past Due (Grace Period)",
            LicenseStatus::Expired => "Expired",
            LicenseStatus::Invalid => "Invalid",
            LicenseStatus::Free => "Free",
        }
    }
}

impl LicenseDisplayInfo {
    pub fn from_current() -> Self {
        match visigrid_license::current_license() {
            Some(lic) if lic.valid => {
                let status = if lic.in_grace_period {
                    LicenseStatus::GracePeriod
                } else {
                    LicenseStatus::Active
                };

                let expires = match lic.expires_at {
                    Some(dt) => dt.format("%Y-%m-%d").to_string(),
                    None => "Never".to_string(),
                };

                LicenseDisplayInfo {
                    edition: lic.edition.to_string(),
                    status,
                    expires,
                    key_id: "embedded".to_string(), // We don't expose key_id in validation currently
                    plan: lic.plan.to_string(),
                }
            }
            Some(lic) => {
                // Invalid license
                LicenseDisplayInfo {
                    edition: "Free".to_string(),
                    status: if lic.error.as_ref().map(|e| e.contains("expired")).unwrap_or(false) {
                        LicenseStatus::Expired
                    } else {
                        LicenseStatus::Invalid
                    },
                    expires: "-".to_string(),
                    key_id: "-".to_string(),
                    plan: "-".to_string(),
                }
            }
            None => {
                LicenseDisplayInfo {
                    edition: "Free".to_string(),
                    status: LicenseStatus::Free,
                    expires: "-".to_string(),
                    key_id: "-".to_string(),
                    plan: "-".to_string(),
                }
            }
        }
    }

    /// Generate diagnostics string for support
    pub fn diagnostics(&self) -> String {
        format!(
            "VisiGrid License Diagnostics\n\
             ----------------------------\n\
             Edition: {}\n\
             Status: {}\n\
             Plan: {}\n\
             Expires: {}\n\
             Key ID: {}\n\
             App Version: {}\n\
             Build: {}",
            self.edition,
            self.status.label(),
            self.plan,
            self.expires,
            self.key_id,
            env!("CARGO_PKG_VERSION"),
            if cfg!(feature = "commercial") { "commercial" } else { "oss" }
        )
    }
}

/// Translate technical license errors to user-friendly messages
pub fn user_friendly_error(error: &str) -> String {
    // JSON parse errors
    if error.contains("Invalid license format") || error.contains("expected") || error.contains("JSON") {
        return "That doesn't look like a license file. Paste the full JSON license.".to_string();
    }

    // Signature errors
    if error.contains("Invalid signature") {
        return "License signature is invalid. Download a fresh license from VisiHub.".to_string();
    }
    if error.contains("signature encoding") || error.contains("signature format") {
        return "License signature is invalid. Download a fresh license from VisiHub.".to_string();
    }

    // Unknown key
    if error.contains("Unknown key ID") {
        return "This license was signed with a newer key. Update VisiGrid to verify it.".to_string();
    }

    // Wrong product
    if error.contains("different product") {
        return "This license is not for VisiGrid.".to_string();
    }

    // Expired
    if error.contains("expired") {
        return "License has expired. Pro features are disabled.".to_string();
    }

    // Revoked
    if error.contains("revoked") {
        return "This license has been revoked. Contact support.".to_string();
    }

    // Not yet valid
    if error.contains("not yet valid") {
        return "This license is not yet valid. Check the start date.".to_string();
    }

    // Algorithm not supported
    if error.contains("Unsupported signature algorithm") {
        return "License uses unsupported format. Update VisiGrid.".to_string();
    }

    // Fallback - don't expose raw technical error
    "License verification failed. Try downloading a fresh license.".to_string()
}

/// Render the license dialog
pub fn render_license_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let error_color = app.token(TokenKey::Error);
    let editor_bg = app.token(TokenKey::EditorBg);
    let success_color = hsla(142.0 / 360.0, 0.71, 0.45, 1.0); // Green (#22c55e)
    let warning_color = hsla(45.0 / 360.0, 0.93, 0.47, 1.0); // Yellow (#eab308)

    let info = LicenseDisplayInfo::from_current();
    let has_license = info.status == LicenseStatus::Active || info.status == LicenseStatus::GracePeriod;

    // Status badge color
    let status_color = match info.status {
        LicenseStatus::Active => success_color,
        LicenseStatus::GracePeriod => warning_color,
        LicenseStatus::Expired | LicenseStatus::Invalid => error_color,
        LicenseStatus::Free => text_muted,
    };

    // Header with close button
    let header = div()
        .flex()
        .items_center()
        .justify_between()
        .child(
            div()
                .text_size(px(14.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_primary)
                .child("License")
        )
        .child(
            div()
                .id("license-close")
                .px_2()
                .cursor_pointer()
                .text_color(text_muted)
                .hover(|s| s.text_color(text_primary))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.hide_license(cx);
                }))
                .child("×")
        );

    // Body content
    let body = div()
        // Status badge at top
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .w(px(10.0))
                        .h(px(10.0))
                        .rounded_full()
                        .bg(status_color)
                )
                .child(
                    div()
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(format!("VisiGrid {}", info.edition))
                )
                .child(
                    div()
                        .px_2()
                        .py_px()
                        .rounded_sm()
                        .bg(status_color.opacity(0.15))
                        .text_xs()
                        .text_color(status_color)
                        .child(info.status.label())
                )
        )
        // Info table
        .child(
            div()
                .p_3()
                .bg(editor_bg)
                .rounded_md()
                .flex()
                .flex_col()
                .gap_1()
                .child(render_info_row("Edition", &info.edition, text_muted, text_primary))
                .child(render_info_row("Status", info.status.label(), text_muted, status_color))
                .when(has_license, |d| {
                    d.child(render_info_row("Plan", &info.plan, text_muted, text_primary))
                })
                .child(render_info_row(
                    "Expires",
                    &info.expires,
                    text_muted,
                    if info.status == LicenseStatus::GracePeriod { warning_color } else { text_primary }
                ))
                .child(
                    div()
                        .mt_2()
                        .pt_2()
                        .border_t_1()
                        .border_color(panel_border)
                        .child(
                            div()
                                .id("copy-diagnostics")
                                .text_xs()
                                .text_color(text_muted)
                                .cursor_pointer()
                                .hover(|s| s.text_color(accent))
                                .on_mouse_down(MouseButton::Left, {
                                    let diagnostics = info.diagnostics();
                                    cx.listener(move |this, _, _window, cx| {
                                        cx.write_to_clipboard(ClipboardItem::new_string(diagnostics.clone()));
                                        this.status_message = Some("Diagnostics copied to clipboard".to_string());
                                        cx.notify();
                                    })
                                })
                                .child("Copy Diagnostics")
                        )
                )
        )
        // Upgrade prompt (only if Free)
        .when(info.status == LicenseStatus::Free, |d| {
            d.child(
                div()
                    .text_xs()
                    .text_color(text_muted)
                    .child("Upgrade to Pro for Lua scripting, large files, and more.")
            )
        })
        // Input section
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(
                    div()
                        .text_xs()
                        .text_color(text_muted)
                        .child(if has_license { "Enter new license" } else { "Enter license key" })
                )
                .child(
                    div()
                        .id("license-input")
                        .h(px(80.0))
                        .w_full()
                        .p_2()
                        .bg(editor_bg)
                        .border_1()
                        .border_color(if app.license_error.is_some() { error_color } else { panel_border })
                        .rounded_md()
                        .text_xs()
                        .font_family("monospace")
                        .text_color(text_primary)
                        .overflow_hidden()
                        .child(
                            if app.license_input.is_empty() {
                                div()
                                    .text_color(text_muted.opacity(0.5))
                                    .child("Paste license JSON here...")
                            } else {
                                div()
                                    .child(truncate_license(&app.license_input, 400))
                            }
                        )
                )
                .when(app.license_error.is_some(), |d| {
                    d.child(
                        div()
                            .p_2()
                            .bg(error_color.opacity(0.1))
                            .rounded_md()
                            .flex()
                            .flex_col()
                            .gap_1()
                            .child(
                                div()
                                    .text_xs()
                                    .font_weight(FontWeight::MEDIUM)
                                    .text_color(error_color)
                                    .child(app.license_error.clone().unwrap_or_default())
                            )
                    )
                })
        );

    // Footer with buttons
    let footer = div()
        .flex()
        .items_center()
        .justify_between()
        .child(
            div()
                .when(has_license, |d| {
                    d.child(
                        div()
                            .id("remove-license")
                            .px_3()
                            .py_1()
                            .rounded_md()
                            .text_xs()
                            .text_color(error_color)
                            .cursor_pointer()
                            .hover(|s| s.bg(error_color.opacity(0.1)))
                            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                this.clear_license(cx);
                            }))
                            .child("Remove License")
                    )
                })
        )
        .child(
            div()
                .flex()
                .gap_2()
                .child(
                    div()
                        .id("license-cancel")
                        .px_4()
                        .py_2()
                        .rounded_md()
                        .border_1()
                        .border_color(panel_border)
                        .text_xs()
                        .text_color(text_muted)
                        .cursor_pointer()
                        .hover(|s| s.bg(panel_border.opacity(0.3)))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.hide_license(cx);
                        }))
                        .child("Cancel")
                )
                .child(
                    div()
                        .id("license-apply")
                        .px_4()
                        .py_2()
                        .rounded_md()
                        .bg(if app.license_input.is_empty() { accent.opacity(0.5) } else { accent })
                        .text_xs()
                        .text_color(rgb(0xffffff))
                        .when(!app.license_input.is_empty(), |d| {
                            d.cursor_pointer()
                                .hover(|s| s.opacity(0.9))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.apply_license(cx);
                                }))
                        })
                        .child("Activate")
                )
        );

    modal_overlay(
        "license-dialog",
        |this, cx| this.hide_license(cx),
        // Wrap DialogFrame in div for keyboard handling
        div()
            .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                let key = &event.keystroke.key;
                match key.as_str() {
                    "escape" => this.hide_license(cx),
                    "enter" => {
                        if !this.license_input.is_empty() {
                            this.apply_license(cx);
                        }
                    }
                    "backspace" => this.license_backspace(cx),
                    _ => {
                        if let Some(c) = event.keystroke.key_char.as_ref().and_then(|s| s.chars().next()) {
                            if !event.keystroke.modifiers.control && !event.keystroke.modifiers.alt {
                                this.license_insert_char(c, cx);
                            }
                        }
                    }
                }
                cx.stop_propagation();
            }))
            .child(
                DialogFrame::new(body, panel_bg, panel_border)
                    .size(DialogSize::Lg)
                    .header(header)
                    .footer(footer)
            ),
        cx,
    )
}

/// Truncate license text for display
fn truncate_license(text: &str, max_chars: usize) -> String {
    if text.len() <= max_chars {
        text.to_string()
    } else {
        format!("{}...", &text[..max_chars])
    }
}

/// Render a label: value info row
fn render_info_row(label: &str, value: &str, label_color: Hsla, value_color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .justify_between()
        .child(
            div()
                .text_xs()
                .text_color(label_color)
                .child(label.to_string())
        )
        .child(
            div()
                .text_xs()
                .font_weight(FontWeight::MEDIUM)
                .text_color(value_color)
                .child(value.to_string())
        )
}
