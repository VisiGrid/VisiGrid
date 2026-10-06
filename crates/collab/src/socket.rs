//! Blocking native socket, driven off the UI thread. All credentials travel
//! in Authorization; neither URLs nor errors expose the bearer token.
use serde_json::Value;
use std::{
    net::{TcpStream, ToSocketAddrs},
    time::Duration,
};
use tungstenite::{client::IntoClientRequest, stream::MaybeTlsStream, Message, WebSocket};

#[derive(Debug, Clone, Copy)]
pub enum ConnectError {
    Fatal(&'static str),
    Transient,
}
impl ConnectError {
    pub fn is_fatal(self) -> bool {
        matches!(self, Self::Fatal(_))
    }
}
impl std::fmt::Display for ConnectError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::Fatal(message) => message,
            Self::Transient => "Could not connect to the collaboration service",
        })
    }
}

fn loopback_host(url: &url::Url) -> bool {
    matches!(url.host(), Some(url::Host::Domain("localhost")))
        || matches!(url.host(), Some(url::Host::Ipv4(ip)) if ip == std::net::Ipv4Addr::LOCALHOST)
        || matches!(url.host(), Some(url::Host::Ipv6(ip)) if ip == std::net::Ipv6Addr::LOCALHOST)
}
pub fn validate_api_origin(value: &str) -> Result<(), ConnectError> {
    validate_endpoint(value, "https", "http").map(|_| ())
}
fn validate_endpoint(value: &str, secure: &str, local: &str) -> Result<url::Url, ConnectError> {
    let url =
        url::Url::parse(value).map_err(|_| ConnectError::Fatal("Invalid collaboration origin"))?;
    if !url.username().is_empty()
        || url.password().is_some()
        || url.host().is_none()
        || !(url.scheme() == secure || (url.scheme() == local && loopback_host(&url)))
    {
        return Err(ConnectError::Fatal(
            "Live collaboration requires TLS except on loopback",
        ));
    }
    Ok(url)
}

#[derive(Default)]
pub struct ReconnectBackoff {
    attempts: u32,
}
impl ReconnectBackoff {
    pub fn reset(&mut self) {
        self.attempts = 0;
    }
    pub fn next_delay(&mut self, jitter: u64) -> Duration {
        let cap = (500u64 * (1u64 << self.attempts.min(6))).min(30_000);
        self.attempts = self.attempts.saturating_add(1);
        Duration::from_millis(cap * 3 / 4 + jitter % (cap / 4 + 1))
    }
}

pub struct LiveSocket(WebSocket<MaybeTlsStream<TcpStream>>);
impl LiveSocket {
    pub fn connect(value: &str, token: &str) -> Result<Self, ConnectError> {
        let url = validate_endpoint(value, "wss", "ws")?;
        let mut request = value
            .into_client_request()
            .map_err(|_| ConnectError::Fatal("Invalid collaboration URL"))?;
        request.headers_mut().insert(
            "Authorization",
            format!("Bearer {token}")
                .parse()
                .map_err(|_| ConnectError::Fatal("Invalid authentication token"))?,
        );
        request.headers_mut().insert(
            "Sec-WebSocket-Protocol",
            "visigrid.collab.v2".parse().unwrap(),
        );
        let addresses = (
            url.host_str().unwrap().trim_matches(['[', ']']),
            url.port_or_known_default().unwrap(),
        )
            .to_socket_addrs()
            .map_err(|_| ConnectError::Transient)?;
        let deadline = std::time::Instant::now() + Duration::from_secs(5);
        let mut stream = None;
        for address in addresses {
            // localhost must not send credentials if a resolver maps it away from loopback.
            if url.scheme() == "ws" && !address.ip().is_loopback() {
                return Err(ConnectError::Fatal(
                    "Cleartext collaboration requires a loopback peer",
                ));
            }
            let Some(remaining) = deadline.checked_duration_since(std::time::Instant::now()) else {
                break;
            };
            if let Ok(connected) = TcpStream::connect_timeout(&address, remaining) {
                stream = Some(connected);
                break;
            }
        }
        let stream = stream.ok_or(ConnectError::Transient)?;
        stream
            .set_read_timeout(Some(Duration::from_secs(5)))
            .map_err(|_| ConnectError::Transient)?;
        stream
            .set_write_timeout(Some(Duration::from_secs(5)))
            .map_err(|_| ConnectError::Transient)?;
        let (mut socket, response) =
            tungstenite::client_tls(request, stream).map_err(|error| match error {
                tungstenite::HandshakeError::Failure(tungstenite::Error::Http(response))
                    if response.status().as_u16() == 401 =>
                {
                    ConnectError::Fatal("Sign in again to resume collaboration")
                }
                tungstenite::HandshakeError::Failure(tungstenite::Error::Http(response))
                    if response.status().is_client_error() =>
                {
                    ConnectError::Fatal("The server refused live workbook access")
                }
                _ => ConnectError::Transient,
            })?;
        if response
            .headers()
            .get("Sec-WebSocket-Protocol")
            .and_then(|v| v.to_str().ok())
            != Some("visigrid.collab.v2")
        {
            return Err(ConnectError::Fatal(
                "The server did not accept the collaboration protocol",
            ));
        }
        match socket.get_mut() {
            MaybeTlsStream::Plain(stream) if url.scheme() == "ws" && loopback_host(&url) => {
                stream.set_read_timeout(Some(Duration::from_millis(50)))
            }
            MaybeTlsStream::Rustls(stream) => stream
                .sock
                .set_read_timeout(Some(Duration::from_millis(50))),
            _ => return Err(ConnectError::Fatal("Unsupported collaboration transport")),
        }
        .map_err(|_| ConnectError::Fatal("Could not configure collaboration transport"))?;
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

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn refuses_cleartext_remote_origins_before_connecting() {
        for value in [
            "http://example.com",
            "http://127.0.0.2",
            "http://localhost.example.com",
            "https://user:secret@example.com",
            "file:///tmp/test",
        ] {
            assert!(validate_api_origin(value).is_err(), "{value}");
        }
        for value in [
            "https://example.com",
            "http://127.0.0.1:8000",
            "http://[::1]:8000",
            "http://localhost:8000",
        ] {
            assert!(validate_api_origin(value).is_ok(), "{value}");
        }
        assert!(matches!(
            LiveSocket::connect("ws://example.com/collab", "secret"),
            Err(ConnectError::Fatal(_))
        ));
    }
    #[test]
    fn backoff_grows_caps_and_resets_with_bounded_jitter() {
        let mut backoff = ReconnectBackoff::default();
        for cap in [500, 1000, 2000, 4000, 8000, 16000, 30000, 30000] {
            let delay = backoff.next_delay(u64::MAX).as_millis();
            assert!((cap * 3 / 4..=cap).contains(&delay));
        }
        backoff.reset();
        assert_eq!(backoff.next_delay(0), Duration::from_millis(375));
    }
}
