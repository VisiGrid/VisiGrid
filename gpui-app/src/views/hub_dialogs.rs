// Hub dialogs: Paste Token and Link to Dataset

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize};
use gpui::{Animation, pulsating_between};

/// Render the paste token dialog (fallback auth flow)
pub fn render_paste_token_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let app_bg = app.token(TokenKey::AppBg);
    let warning_color = hsla(0.08, 0.9, 0.55, 1.0); // Orange for security warning

    let token_value = app.hub_token_input.clone();
    let has_token = !token_value.is_empty();

    // Body content
    let body = div()
        .flex()
        .flex_col()
        .gap_4()
        // Title
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_size(px(18.0))
                        .font_weight(FontWeight::BOLD)
                        .text_color(text_primary)
                        .child("Sign in")
                )
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(text_muted)
                        .child("A browser window should have opened. Complete sign in there, then paste your token below.")
                )
        )
        // Security warning
        .child(
            div()
                .px_3()
                .py_2()
                .bg(hsla(0.08, 0.9, 0.55, 0.15)) // Orange tint background
                .border_1()
                .border_color(warning_color)
                .rounded_md()
                .text_size(px(12.0))
                .text_color(warning_color)
                .child("Only paste tokens from visigrid.app")
        )
        // Token input
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_muted)
                        .child("Device Token")
                )
                .child(
                    div()
                        .w_full()
                        .h(px(36.0))
                        .px_3()
                        .bg(app_bg)
                        .border_1()
                        .border_color(panel_border)
                        .rounded_md()
                        .flex()
                        .items_center()
                        .text_color(text_primary)
                        .text_size(px(13.0))
                        .child(if token_value.is_empty() {
                            div().text_color(text_muted).child("Paste token here...|")
                        } else {
                            // Mask entire token like a password field
                            div().child(format!("{}|", mask_token_full(&token_value)))
                        })
                )
        )
        // Help text
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("If the browser didn't open, visit visigrid.app/desktop/authorize")
        );

    // Footer with buttons
    let footer = div()
        .flex()
        .gap_2()
        .justify_end()
        .child(
            Button::new("cancel-sign-in", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.hub_cancel_sign_in(cx);
                }))
        )
        .child(
            Button::new("confirm-sign-in", "Sign In")
                .disabled(!has_token)
                .primary(accent, hsla(0.0, 0.0, 1.0, 1.0))
                .when(has_token, |d| {
                    d.on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                        cx.stop_propagation();
                        this.hub_complete_sign_in(cx);
                    }))
                })
        );

    modal_overlay(
        "paste-token-dialog",
        |this, cx| this.hub_cancel_sign_in(cx),
        DialogFrame::new(body, panel_bg, panel_border)
            .width(px(420.0))
            .footer(footer),
        cx,
    )
}

/// Render the link to dataset dialog
pub fn render_link_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let app_bg = app.token(TokenKey::AppBg);

    let loading = app.hub_link_loading;
    let repos = app.hub_repos.clone();
    let selected_repo = app.hub_selected_repo;
    let datasets = app.hub_datasets.clone();
    let selected_dataset = app.hub_selected_dataset;
    let new_dataset_name = app.hub_new_dataset_name.clone();

    let can_link = selected_repo.is_some() && selected_dataset.is_some();
    let can_create = selected_repo.is_some() && !new_dataset_name.trim().is_empty();

    // Body content
    let body = div()
        .flex()
        .flex_col()
        .gap_4()
        // Title
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_size(px(18.0))
                        .font_weight(FontWeight::BOLD)
                        .text_color(text_primary)
                        .child("Link to Repository")
                )
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(text_muted)
                        .child("Select a repository and dataset to link this workbook.")
                )
        )
        // Loading indicator
        .when(loading, |el| {
            el.child(
                div()
                    .text_size(px(13.0))
                    .text_color(text_muted)
                    .child("Loading...")
            )
        })
        // Repository list
        .when(!loading && !repos.is_empty(), |el| {
            el.child(render_repo_list(&repos, selected_repo, panel_bg, panel_border, text_primary, text_muted, selection_bg, cx))
        })
        // Dataset list (when repo selected)
        .when(selected_repo.is_some() && !loading, |el| {
            el.child(render_dataset_list(&datasets, selected_dataset, panel_bg, panel_border, text_primary, text_muted, selection_bg, cx))
        })
        // Create new dataset
        .when(selected_repo.is_some() && !loading, |el| {
            el.child(render_create_dataset(&new_dataset_name, can_create, panel_border, text_muted, text_primary, app_bg, accent, cx))
        })
        // Mode info
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("Mode: Pull only (receive updates)")
        );

    // Footer with buttons
    let footer = render_link_buttons(can_link, panel_border, text_muted, accent, cx);

    modal_overlay(
        "link-dialog",
        |this, cx| this.hub_cancel_link(cx),
        DialogFrame::new(body, panel_bg, panel_border)
            .width(px(480.0))
            .max_height(px(500.0))
            .footer(footer),
        cx,
    )
}

fn render_repo_list(
    repos: &[crate::hub::RepoInfo],
    selected_repo: Option<usize>,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child("Repository")
        )
        .child(
            div()
                .max_h(px(120.0))
                .overflow_hidden()
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .children(
                    repos.iter().enumerate().map(|(idx, repo)| {
                        let is_selected = selected_repo == Some(idx);
                        let bg = if is_selected { selection_bg } else { panel_bg };
                        let text_col = if is_selected { text_primary } else { text_muted };
                        let display = format!("{}/{}", repo.owner, repo.slug);

                        div()
                            .id(ElementId::Name(format!("repo-{}", idx).into()))
                            .px_3()
                            .py_2()
                            .bg(bg)
                            .text_color(text_col)
                            .text_size(px(13.0))
                            .cursor_pointer()
                            .hover(move |s| s.bg(selection_bg))
                            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                cx.stop_propagation();
                                this.hub_select_repo(idx, cx);
                            }))
                            .child(display)
                    })
                )
        )
}

fn render_dataset_list(
    datasets: &[crate::hub::DatasetInfo],
    selected_dataset: Option<usize>,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child("Dataset")
        )
        .when(datasets.is_empty(), |d| {
            d.child(
                div()
                    .text_size(px(13.0))
                    .text_color(text_muted)
                    .italic()
                    .child("No datasets yet. Create one below.")
            )
        })
        .when(!datasets.is_empty(), |d| {
            d.child(
                div()
                    .max_h(px(120.0))
                    .overflow_hidden()
                    .border_1()
                    .border_color(panel_border)
                    .rounded_md()
                    .children(
                        datasets.iter().enumerate().map(|(idx, dataset)| {
                            let is_selected = selected_dataset == Some(idx);
                            let bg = if is_selected { selection_bg } else { panel_bg };
                            let text_col = if is_selected { text_primary } else { text_muted };
                            let name = dataset.name.clone();

                            div()
                                .id(ElementId::Name(format!("dataset-{}", idx).into()))
                                .px_3()
                                .py_2()
                                .bg(bg)
                                .text_color(text_col)
                                .text_size(px(13.0))
                                .cursor_pointer()
                                .hover(move |s| s.bg(selection_bg))
                                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                    cx.stop_propagation();
                                    this.hub_select_dataset(idx, cx);
                                }))
                                .child(name)
                        })
                    )
            )
        })
}

fn render_create_dataset(
    new_dataset_name: &str,
    can_create: bool,
    panel_border: Hsla,
    text_muted: Hsla,
    text_primary: Hsla,
    app_bg: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let name = new_dataset_name.to_string();

    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child("Or create new dataset")
        )
        .child(
            div()
                .flex()
                .gap_2()
                .child(
                    div()
                        .id("dataset-name-input")
                        .flex_1()
                        .h(px(32.0))
                        .px_3()
                        .bg(app_bg)
                        .border_1()
                        .border_color(panel_border)
                        .rounded_md()
                        .flex()
                        .items_center()
                        .text_color(text_primary)
                        .text_size(px(13.0))
                        .cursor_text()
                        .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                            cx.stop_propagation();
                            let key = event.keystroke.key.as_str();
                            if key == "backspace" {
                                this.hub_dataset_backspace(cx);
                            } else if key == "enter" && !this.hub_new_dataset_name.trim().is_empty() {
                                this.hub_create_and_link(cx);
                            } else if let Some(key_char) = &event.keystroke.key_char {
                                // Only insert if no modifiers (except shift for uppercase)
                                if !event.keystroke.modifiers.control
                                    && !event.keystroke.modifiers.alt
                                    && !event.keystroke.modifiers.platform
                                {
                                    for c in key_char.chars() {
                                        this.hub_dataset_insert_char(c, cx);
                                    }
                                }
                            }
                        }))
                        .child(if name.is_empty() {
                            div().text_color(text_muted).child("Dataset name...")
                        } else {
                            div().child(name.clone())
                        })
                        // Blinking cursor
                        .child(
                            div()
                                .w(px(1.0))
                                .h(px(14.0))
                                .bg(text_primary)
                                .with_animation(
                                    "cursor-blink",
                                    Animation::new(std::time::Duration::from_millis(530))
                                        .repeat()
                                        .with_easing(pulsating_between(0.0, 1.0)),
                                    |div, delta| div.opacity(if delta > 0.5 { 0.0 } else { 1.0 }),
                                )
                        )
                )
                .child(
                    div()
                        .id("create-dataset-btn")
                        .px_4()
                        .py_2()
                        .bg(if can_create { accent } else { hsla(0.0, 0.0, 0.3, 1.0) })
                        .rounded_md()
                        .text_color(hsla(0.0, 0.0, 1.0, 1.0))
                        .text_size(px(13.0))
                        .cursor(if can_create { CursorStyle::PointingHand } else { CursorStyle::Arrow })
                        .when(can_create, |d| {
                            d.on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                cx.stop_propagation();
                                this.hub_create_and_link(cx);
                            }))
                        })
                        .child("Create")
                )
        )
}

fn render_link_buttons(
    can_link: bool,
    panel_border: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .gap_2()
        .justify_end()
        .child(
            Button::new("cancel-link", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.hub_cancel_link(cx);
                }))
        )
        .child(
            Button::new("confirm-link", "Link")
                .disabled(!can_link)
                .primary(accent, hsla(0.0, 0.0, 1.0, 1.0))
                .when(can_link, |d| {
                    d.on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                        cx.stop_propagation();
                        this.hub_confirm_link(cx);
                    }))
                })
        )
}

/// Mask a token for display (show first 8 chars + ...)
#[allow(dead_code)]
fn mask_token(token: &str) -> String {
    if token.len() <= 8 {
        token.to_string()
    } else {
        format!("{}...", &token[..8])
    }
}

/// Mask entire token like a password field (all dots)
fn mask_token_full(token: &str) -> String {
    "•".repeat(token.len().min(32)) // Cap at 32 dots for visual sanity
}

/// Render the publish confirmation dialog (shown when diverged)
pub fn render_publish_confirm_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let warning_color = hsla(0.08, 0.9, 0.55, 1.0); // Orange
    let danger_color = hsla(0.0, 0.7, 0.5, 1.0); // Red-ish

    // Body content
    let body = div()
        .flex()
        .flex_col()
        .gap_4()
        // Warning icon and title
        .child(
            div()
                .flex()
                .items_center()
                .gap_3()
                .child(
                    div()
                        .text_size(px(24.0))
                        .text_color(warning_color)
                        .child("!")
                )
                .child(
                    div()
                        .text_size(px(18.0))
                        .font_weight(FontWeight::BOLD)
                        .text_color(text_primary)
                        .child("Remote has changed")
                )
        )
        // Body text
        .child(
            div()
                .text_size(px(13.0))
                .text_color(text_muted)
                .child("The remote has a newer version of this workbook. Publishing now will replace the remote version.")
        );

    // Footer with buttons
    let footer = div()
        .flex()
        .gap_2()
        .justify_end()
        .child(
            Button::new("cancel-publish", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.hub_cancel_publish_confirm(cx);
                }))
        )
        .child(
            div()
                .id("open-remote-copy")
                .px_4()
                .py_2()
                .bg(panel_border)
                .rounded_md()
                .text_color(text_primary)
                .text_size(px(13.0))
                .cursor_pointer()
                .hover(move |s| s.bg(hsla(0.0, 0.0, 0.4, 1.0)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.hub_cancel_publish_confirm(cx);
                    this.hub_open_remote_as_copy(cx);
                }))
                .child("Open Remote as Copy")
        )
        .child(
            Button::new("publish-anyway", "Publish Anyway")
                .primary(danger_color, hsla(0.0, 0.0, 1.0, 1.0))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.hub_confirm_publish_anyway(cx);
                }))
        );

    modal_overlay(
        "publish-confirm-dialog",
        |this, cx| this.hub_cancel_publish_confirm(cx),
        DialogFrame::new(body, panel_bg, panel_border)
            .width(px(400.0)),
        cx,
    )
}
