// Backend-agnostic session policy. Authentication providers may emit refresh/recovery events;
// the UI receives a stable session transition instead of provider-specific navigation behavior.
export function createSessionController(auth) {
  let subscription = null;
  let currentSession = undefined;
  let generation = 0;
  return {
    async start(onChange) {
      const myGeneration = ++generation;
      const { data: initial, error } = await auth.getSession();
      if (myGeneration !== generation) return;
      if (error) throw error;
      currentSession = initial.session || null;
      onChange({ session: currentSession, entered: false, reason: "startup" });
      const result = auth.onAuthStateChange((event, session) => {
        if (myGeneration !== generation) return;
        const previous = currentSession;
        currentSession = session || null;
        const entered = !previous && !!currentSession && event === "SIGNED_IN";
        onChange({ session: currentSession, entered, reason: event });
      });
      subscription = result?.data?.subscription || null;
    },
    stop() {
      generation++;
      subscription?.unsubscribe?.();
      subscription = null;
    },
    get current() { return currentSession; },
  };
}
