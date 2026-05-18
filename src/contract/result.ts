export type Outcome<T> = Ok<T> | Fail;

export interface Ok<T> {
  ok: true;
  data: T;
  warns: string[];
}

export interface Fail {
  ok: false;
  error: string;
  warns: string[];
}

export function succeed<T>(data: T, warns: string[] = []): Ok<T> {
  return { ok: true, data, warns };
}

export function fail(error: string, warns: string[] = []): Fail {
  return { ok: false, error, warns };
}
