"use strict";

const { newInternalId } = require("../utils/fileUtils");
const config = require("../config");

/**
 * Simple in-memory session store.
 *
 * Holds the state between "/preview" (analyze + AI plan) and "/apply"
 * (execute + verify + download) so the user can review planned changes
 * before anything is written. Good enough for a single-process deployment;
 * swap for Redis if you need multi-instance scaling.
 */
const sessions = new Map();

function create(data) {
  const id = newInternalId();
  sessions.set(id, {
    id,
    createdAt: Date.now(),
    ...data,
  });
  return id;
}

function get(id) {
  const session = sessions.get(id);
  if (!session) return null;
  if (Date.now() - session.createdAt > config.fileTtlMs) {
    sessions.delete(id);
    return null;
  }
  return session;
}

function update(id, patch) {
  const session = sessions.get(id);
  if (!session) return null;
  Object.assign(session, patch);
  return session;
}

function remove(id) {
  sessions.delete(id);
}

function sweepExpired() {
  const now = Date.now();
  for (const [id, session] of sessions.entries()) {
    if (now - session.createdAt > config.fileTtlMs) {
      sessions.delete(id);
    }
  }
}

module.exports = { create, get, update, remove, sweepExpired };
