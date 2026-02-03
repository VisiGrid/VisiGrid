//! Event broadcasting for session server.
//!
//! The event system allows clients to subscribe to workbook changes and receive
//! real-time notifications when cells are modified.
//!
//! Design:
//! - Each connection has its own subscription set (topics)
//! - GUI thread broadcasts events through a channel
//! - Connection threads poll for events and forward to subscribed clients
//! - Cell changes are coalesced into ranges before broadcasting (efficient)

use std::collections::HashSet;
use std::sync::mpsc::{self, Receiver, Sender, TryRecvError};

use super::protocol::{CellRange, EventMessage, EventPayload, ServerMessage};

/// Topic for cell change events.
pub const TOPIC_CELLS: &str = "cells";

/// All valid topics.
pub const VALID_TOPICS: &[&str] = &[TOPIC_CELLS];

/// Event sent from GUI thread to connection handlers.
/// Ranges are pre-coalesced at broadcast time for efficiency.
#[derive(Debug, Clone)]
pub struct BroadcastEvent {
    /// Revision that produced this event.
    pub revision: u64,
    /// Coalesced ranges covering all changed cells.
    pub ranges: Vec<CellRange>,
}

/// Handle for broadcasting events from the GUI thread.
/// Cloneable - can be used from multiple places.
#[derive(Clone)]
pub struct EventBroadcaster {
    /// Sender to the broadcast channel.
    /// None if no listeners are connected.
    tx: Option<Sender<BroadcastEvent>>,
}

impl EventBroadcaster {
    /// Create a new broadcaster (initially disconnected).
    pub fn new() -> Self {
        Self { tx: None }
    }

    /// Create a broadcaster with a connected channel.
    pub fn with_channel(tx: Sender<BroadcastEvent>) -> Self {
        Self { tx: Some(tx) }
    }

    /// Broadcast a cell change event with pre-coalesced ranges.
    /// No-op if no listeners are connected.
    pub fn broadcast_ranges(&self, revision: u64, ranges: Vec<CellRange>) {
        if let Some(tx) = &self.tx {
            let _ = tx.send(BroadcastEvent { revision, ranges });
        }
    }

    /// Check if broadcaster has any listeners.
    pub fn has_listeners(&self) -> bool {
        self.tx.is_some()
    }
}

impl Default for EventBroadcaster {
    fn default() -> Self {
        Self::new()
    }
}

/// Subscription state for a single connection.
pub struct ConnectionSubscriptions {
    /// Subscribed topics.
    topics: HashSet<String>,
    /// Receiver for broadcast events.
    event_rx: Receiver<BroadcastEvent>,
}

impl ConnectionSubscriptions {
    /// Create subscription state with a broadcast receiver.
    pub fn new(event_rx: Receiver<BroadcastEvent>) -> Self {
        Self {
            topics: HashSet::new(),
            event_rx,
        }
    }

    /// Subscribe to topics. Returns list of successfully subscribed topics.
    pub fn subscribe(&mut self, topics: &[String]) -> Vec<String> {
        let mut subscribed = Vec::new();
        for topic in topics {
            if VALID_TOPICS.contains(&topic.as_str()) {
                if self.topics.insert(topic.clone()) {
                    subscribed.push(topic.clone());
                }
            }
        }
        subscribed
    }

    /// Unsubscribe from topics. Returns list of successfully unsubscribed topics.
    pub fn unsubscribe(&mut self, topics: &[String]) -> Vec<String> {
        let mut unsubscribed = Vec::new();
        for topic in topics {
            if self.topics.remove(topic) {
                unsubscribed.push(topic.clone());
            }
        }
        unsubscribed
    }

    /// Check if subscribed to a topic.
    pub fn is_subscribed(&self, topic: &str) -> bool {
        self.topics.contains(topic)
    }

    /// Get all subscribed topics.
    pub fn subscribed_topics(&self) -> Vec<String> {
        self.topics.iter().cloned().collect()
    }

    /// Poll for pending events. Returns server messages to send.
    /// Non-blocking - returns empty vec if no events pending.
    pub fn poll_events(&self) -> Vec<ServerMessage> {
        let mut messages = Vec::new();

        // Only process if subscribed to cells
        if !self.is_subscribed(TOPIC_CELLS) {
            // Drain the channel but don't produce messages
            while self.event_rx.try_recv().is_ok() {}
            return messages;
        }

        // Collect all pending events
        loop {
            match self.event_rx.try_recv() {
                Ok(event) => {
                    if !event.ranges.is_empty() {
                        messages.push(ServerMessage::Event(EventMessage {
                            topic: TOPIC_CELLS.to_string(),
                            revision: event.revision,
                            payload: EventPayload::CellsChanged { ranges: event.ranges },
                        }));
                    }
                }
                Err(TryRecvError::Empty) => break,
                Err(TryRecvError::Disconnected) => break,
            }
        }

        messages
    }
}

/// Factory for creating per-connection subscriptions.
/// Each connection gets its own broadcast receiver.
pub struct SubscriptionFactory {
    /// Template broadcaster for creating receivers.
    /// When a connection subscribes, it gets a clone of this receiver.
    broadcast_tx: Sender<BroadcastEvent>,
    broadcast_rx: Option<Receiver<BroadcastEvent>>,
}

impl SubscriptionFactory {
    /// Create a new subscription factory.
    /// Returns the factory and a broadcaster for the GUI thread.
    pub fn new() -> (Self, EventBroadcaster) {
        let (tx, rx) = mpsc::channel();
        let factory = Self {
            broadcast_tx: tx.clone(),
            broadcast_rx: Some(rx),
        };
        let broadcaster = EventBroadcaster::with_channel(tx);
        (factory, broadcaster)
    }

    /// Create a new connection subscription.
    /// Note: This is a simplified implementation - all connections share the
    /// same channel. In production, we'd use a proper broadcast channel.
    pub fn create_subscription(&self) -> ConnectionSubscriptions {
        // Create a new channel for this connection
        let (tx, rx) = mpsc::channel();
        // Note: We'd need to register this tx somewhere to receive broadcasts.
        // For now, create independent channel (will be wired in server.rs)
        ConnectionSubscriptions::new(rx)
    }

    /// Get the broadcast sender for creating per-connection subscriptions.
    pub fn broadcast_sender(&self) -> Sender<BroadcastEvent> {
        self.broadcast_tx.clone()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_subscribe_unsubscribe() {
        let (_tx, rx) = mpsc::channel();
        let mut subs = ConnectionSubscriptions::new(rx);

        // Subscribe to cells
        let subscribed = subs.subscribe(&["cells".to_string()]);
        assert_eq!(subscribed, vec!["cells"]);
        assert!(subs.is_subscribed("cells"));

        // Duplicate subscribe returns empty
        let subscribed = subs.subscribe(&["cells".to_string()]);
        assert!(subscribed.is_empty());

        // Invalid topic ignored
        let subscribed = subs.subscribe(&["invalid".to_string()]);
        assert!(subscribed.is_empty());

        // Unsubscribe
        let unsubscribed = subs.unsubscribe(&["cells".to_string()]);
        assert_eq!(unsubscribed, vec!["cells"]);
        assert!(!subs.is_subscribed("cells"));

        // Duplicate unsubscribe returns empty
        let unsubscribed = subs.unsubscribe(&["cells".to_string()]);
        assert!(unsubscribed.is_empty());
    }

    #[test]
    fn test_poll_events_with_subscription() {
        let (tx, rx) = mpsc::channel();
        let mut subs = ConnectionSubscriptions::new(rx);

        // Subscribe to cells
        subs.subscribe(&["cells".to_string()]);

        // Send an event with pre-coalesced ranges
        tx.send(BroadcastEvent {
            revision: 5,
            ranges: vec![CellRange {
                sheet: 0,
                r1: 0,
                c1: 0,
                r2: 0,
                c2: 0,
            }],
        })
        .unwrap();

        // Poll should return the event
        let messages = subs.poll_events();
        assert_eq!(messages.len(), 1);

        if let ServerMessage::Event(event) = &messages[0] {
            assert_eq!(event.topic, "cells");
            assert_eq!(event.revision, 5);
            if let EventPayload::CellsChanged { ranges } = &event.payload {
                assert_eq!(ranges.len(), 1);
                assert_eq!(ranges[0].r1, 0);
                assert_eq!(ranges[0].c1, 0);
            } else {
                panic!("Expected CellsChanged payload");
            }
        } else {
            panic!("Expected Event message");
        }
    }

    #[test]
    fn test_poll_events_without_subscription() {
        let (tx, rx) = mpsc::channel();
        let subs = ConnectionSubscriptions::new(rx);

        // Don't subscribe

        // Send an event
        tx.send(BroadcastEvent {
            revision: 5,
            ranges: vec![CellRange {
                sheet: 0,
                r1: 0,
                c1: 0,
                r2: 0,
                c2: 0,
            }],
        })
        .unwrap();

        // Poll should drain but return empty
        let messages = subs.poll_events();
        assert!(messages.is_empty());
    }

    #[test]
    fn test_broadcaster() {
        let (tx, rx) = mpsc::channel::<BroadcastEvent>();
        let broadcaster = EventBroadcaster::with_channel(tx);

        assert!(broadcaster.has_listeners());

        broadcaster.broadcast_ranges(
            10,
            vec![CellRange {
                sheet: 0,
                r1: 1,
                c1: 2,
                r2: 1,
                c2: 2,
            }],
        );

        let event = rx.recv().unwrap();
        assert_eq!(event.revision, 10);
        assert_eq!(event.ranges.len(), 1);
        assert_eq!(event.ranges[0].r1, 1);
        assert_eq!(event.ranges[0].c1, 2);
    }

    #[test]
    fn test_broadcaster_disconnected() {
        let broadcaster = EventBroadcaster::new();
        assert!(!broadcaster.has_listeners());

        // Should not panic
        broadcaster.broadcast_ranges(
            10,
            vec![CellRange {
                sheet: 0,
                r1: 1,
                c1: 2,
                r2: 1,
                c2: 2,
            }],
        );
    }
}
