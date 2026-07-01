// Replaces the Claude artifact's window.storage with a browser-native one.
// Data persists per browser via localStorage. Same async shape the app expects.
(function installStorage() {
  if (typeof window === "undefined") return;
  if (window.storage && window.storage.__seazen) return;

  const ls = () => {
    try { return window.localStorage; } catch { return null; }
  };

  window.storage = {
    __seazen: true,
    async get(key) {
      const store = ls();
      const value = store ? store.getItem(key) : null;
      return { key, value }; // value is null when the key is missing
    },
    async set(key, value) {
      const store = ls();
      if (store) store.setItem(key, value);
      return { key, value };
    },
    async delete(key) {
      const store = ls();
      if (store) store.removeItem(key);
      return { key, deleted: true };
    },
    async list(prefix = "") {
      const store = ls();
      const keys = [];
      if (store) {
        for (let i = 0; i < store.length; i++) {
          const k = store.key(i);
          if (!prefix || (k && k.startsWith(prefix))) keys.push(k);
        }
      }
      return { keys, prefix };
    },
  };
})();
