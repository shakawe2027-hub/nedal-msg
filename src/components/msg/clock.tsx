import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { getPublicInfo } from "@/lib/msg.functions";
export function ClockBar() {
 const [now,setNow] = useState<Date | null>(null);
 const q = useQuery({ queryKey: ["institution-info"], queryFn: () => getPublicInfo(), staleTime: 30000 });
 useEffect(() => { setNow(new Date()); const id=setInterval(() => setNow(new Date()),1000); return () => clearInterval(id); }, []);
 return <div className="clock-bar no-print"><b>{q.data?.institution}</b><span>العام الدراسي: {q.data?.year ?? "…"}</span><time>{now?.toLocaleString("ar-PS",{weekday:"long",year:"numeric",month:"long",day:"numeric",hour:"2-digit",minute:"2-digit",second:"2-digit"})}</time></div>;
}
