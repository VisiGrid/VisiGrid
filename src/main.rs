mod app;
mod config;
mod core;
mod engine;
mod io;
mod ui;

use app::App;

fn main() -> iced::Result {
    iced::application("VisiGrid", App::update, App::view)
        .subscription(App::subscription)
        .run()
}
