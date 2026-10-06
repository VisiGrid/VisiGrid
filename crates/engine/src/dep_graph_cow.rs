//! Sharded copy-on-write maps for dependency snapshots. A candidate changing a
//! few edges copies only their hash buckets, not every formula in the workbook.
use rustc_hash::{FxHashMap, FxHasher};
use std::{
    hash::{Hash, Hasher},
    sync::Arc,
};
const SHARDS: usize = 256;
#[derive(Clone, Debug)]
pub(super) struct CowMap<K, V> {
    shards: Vec<Arc<FxHashMap<K, V>>>,
}
impl<K, V> Default for CowMap<K, V> {
    fn default() -> Self {
        Self { shards: Vec::new() }
    }
}
impl<K: Hash + Eq + Clone, V: Clone> CowMap<K, V> {
    fn index(key: &K) -> usize {
        let mut h = FxHasher::default();
        key.hash(&mut h);
        (h.finish() as usize) % SHARDS
    }
    fn shard_mut(&mut self, key: &K) -> &mut FxHashMap<K, V> {
        if self.shards.is_empty() {
            self.shards.resize(SHARDS, Arc::new(FxHashMap::default()));
        }
        Arc::make_mut(&mut self.shards[Self::index(key)])
    }
    pub fn get(&self, key: &K) -> Option<&V> {
        self.shards.get(Self::index(key))?.get(key)
    }
    pub fn get_mut(&mut self, key: &K) -> Option<&mut V> {
        if self.get(key).is_none() {
            return None;
        }
        self.shard_mut(key).get_mut(key)
    }
    pub fn contains_key(&self, key: &K) -> bool {
        self.get(key).is_some()
    }
    pub fn insert(&mut self, key: K, value: V) -> Option<V> {
        self.shard_mut(&key).insert(key, value)
    }
    pub fn remove(&mut self, key: &K) -> Option<V> {
        if self.get(key).is_none() {
            return None;
        }
        self.shard_mut(key).remove(key)
    }
    pub fn entry(&mut self, key: K) -> std::collections::hash_map::Entry<'_, K, V> {
        self.shard_mut(&key).entry(key)
    }
    pub fn iter(&self) -> impl Iterator<Item = (&K, &V)> {
        self.shards.iter().flat_map(|s| s.iter())
    }
    pub fn keys(&self) -> impl Iterator<Item = &K> {
        self.iter().map(|(k, _)| k)
    }
    pub fn values_mut(&mut self) -> impl Iterator<Item = &mut V> {
        self.shards
            .iter_mut()
            .flat_map(|s| Arc::make_mut(s).values_mut())
    }
    pub fn retain(&mut self, mut f: impl FnMut(&K, &mut V) -> bool) {
        for s in &mut self.shards {
            Arc::make_mut(s).retain(&mut f);
        }
    }
    pub fn len(&self) -> usize {
        self.shards.iter().map(|s| s.len()).sum()
    }
    pub fn is_empty(&self) -> bool {
        self.shards.iter().all(|s| s.is_empty())
    }
}
impl<'a, K: Hash + Eq + Clone, V: Clone> IntoIterator for &'a CowMap<K, V> {
    type Item = (&'a K, &'a V);
    type IntoIter = Box<dyn Iterator<Item = Self::Item> + 'a>;
    fn into_iter(self) -> Self::IntoIter {
        Box::new(self.iter())
    }
}
