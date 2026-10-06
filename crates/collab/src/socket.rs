//! Blocking native socket, driven off the UI thread. All credentials travel
//! in Authorization; neither URLs nor errors expose the bearer token.
use serde_json::Value;
use std::{net::TcpStream, time::Duration};
use tungstenite::{client::IntoClientRequest, stream::MaybeTlsStream, Message, WebSocket};

pub struct LiveSocket(WebSocket<MaybeTlsStream<TcpStream>>);
impl LiveSocket {
    pub fn connect(url: &str, token: &str) -> Result<Self, String> {
        let mut request = url
            .into_client_request()
            .map_err(|_| "Invalid collaboration URL")?;
        request.headers_mut().insert(
            "Authorization",
            format!("Bearer {token}")
                .parse()
                .map_err(|_| "Invalid authentication token")?,
        );
        request.headers_mut().insert(
            "Sec-WebSocket-Protocol",
            "visigrid.collab.v2".parse().unwrap(),
        );
        let (mut socket, response) =
            tungstenite::connect(request).map_err(|error| match error {
                tungstenite::Error::Http(response) if response.status().as_u16() == 401 => {
                String::from("Sign in again to resume collaboration")
                }
                _ => "Could not connect to the collaboration service".into(),
            })?;
        if response
            .headers()
            .get("Sec-WebSocket-Protocol")
            .and_then(|v| v.to_str().ok())
            != Some("visigrid.collab.v2")
        {
            return Err("The server did not accept the collaboration protocol".into());
        }
        match socket.get_mut() {
            MaybeTlsStream::Plain(stream) => {
                stream.set_read_timeout(Some(Duration::from_millis(50)))
            }
            MaybeTlsStream::Rustls(stream) => stream
                .sock
                .set_read_timeout(Some(Duration::from_millis(50))),
            _ => return Err("Unsupported collaboration transport".into()),
        }
        .map_err(|_| "Could not configure collaboration transport")?;
        Ok(Self(socket))
    }
    pub fn send(&mut self, frame: &Value) -> Result<(), String> {
        self.0
            .send(Message::Text(frame.to_string()))
            .map_err(|_| "Collaboration connection closed".into())
    }
    pub fn read(&mut self) -> Result<Option<Value>, String> {
        match self.0.read() {
            Ok(Message::Text(text)) => serde_json::from_str(&text)
                .map(Some)
                .map_err(|_| "Invalid collaboration frame".into()),
            Ok(Message::Ping(_)) | Ok(Message::Pong(_)) => {
                let _ = self.0.flush();
                Ok(None)
            }
            Ok(Message::Close(_)) => Err("Collaboration connection closed".into()),
            Ok(_) => Err("Unsupported collaboration frame".into()),
            Err(tungstenite::Error::Io(error))
                if matches!(
                    error.kind(),
                    std::io::ErrorKind::WouldBlock | std::io::ErrorKind::TimedOut
                ) =>
            {
                Ok(None)
            }
            Err(_) => Err("Collaboration connection closed".into()),
        }
    }
}
