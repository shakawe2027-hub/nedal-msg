import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { getMsgData } from "./msg.functions";

export type Session = { code: string; role: "admin" | "user" };
const KEY = "nedal-msg-session";

export function saveSession(s: Session | null) {
  if (s) sessionStorage.setItem(KEY, JSON.stringify(s));
  else sessionStorage.removeItem(KEY);
  window.dispatchEvent(new Event("nedal-session"));
}

/** undefined = not read yet (SSR / before hydration), null = signed out. */
export function useSession() {
  const [s, set] = useState<Session | null | undefined>(undefined);
  useEffect(() => {
    const read = () => {
      try {
        set(JSON.parse(sessionStorage.getItem(KEY) ?? "null"));
      } catch {
        set(null);
      }
    };
    read();
    window.addEventListener("nedal-session", read);
    return () => window.removeEventListener("nedal-session", read);
  }, []);
  return s;
}

export function useMsgData(code: string | undefined) {
  return useQuery({
    queryKey: ["msg", code],
    queryFn: () => getMsgData({ data: { code: code! } }),
    enabled: !!code,
    staleTime: 30_000,
  });
}
