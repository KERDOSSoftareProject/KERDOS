// Compatibility bridge for screens not yet migrated to named KERDOS services.
// Provider construction and configuration live exclusively under backend/.
// Delete this module after commands.table has no callers.
import { backend, backendInfo } from "./backend/index.js";
import { createSessionController } from "./session.js";

export const data = {
  from: backend.commands.table,
};

// Reports which backend a running build is actually pointed at, so a
// deployment can be identified without reading the bundle.
// Session policy lives with the backend boundary, since "how do we know
// who is signed in" is a property of the backend, not of any screen.
// Swapping backends swaps this too, and the UI is unaffected.
export const sessionController = createSessionController(backend.session);

export { backend, backendInfo };
