import fs from "node:fs/promises";
import path from "node:path";
import type { PersistedState } from "./types.js";

const DEFAULT_STATE: PersistedState = {
  records: [],
  sentEventIds: [],
  updatedAt: null
};

export class StateStore {
  constructor(private readonly stateDirectory: string) {}

  private get stateFilePath(): string {
    return path.join(this.stateDirectory, "current-state.json");
  }

  async ensureDirectory(): Promise<void> {
    await fs.mkdir(this.stateDirectory, { recursive: true });
  }

  async load(): Promise<PersistedState> {
    await this.ensureDirectory();

    try {
      const raw = await fs.readFile(this.stateFilePath, "utf-8");
      return JSON.parse(raw) as PersistedState;
    } catch (error) {
      const nodeError = error as NodeJS.ErrnoException;
      if (nodeError.code === "ENOENT") {
        return DEFAULT_STATE;
      }
      throw error;
    }
  }

  async save(state: PersistedState): Promise<void> {
    await this.ensureDirectory();
    await fs.writeFile(this.stateFilePath, JSON.stringify(state, null, 2), "utf-8");
  }
}
