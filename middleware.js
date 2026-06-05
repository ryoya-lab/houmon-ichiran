export const config = { matcher: "/(.*)" };

export default function middleware(request) {
  const auth = request.headers.get("authorization");
  const USER = process.env.BASIC_AUTH_USER || "tassha";
  const PASS = process.env.BASIC_AUTH_PASS || "";

  if (auth) {
    const [scheme, encoded] = auth.split(" ");
    if (scheme === "Basic" && encoded) {
      const [user, pass] = atob(encoded).split(":");
      if (user === USER && pass === PASS && PASS !== "") {
        return; // 認証OK → 静的ファイルへ
      }
    }
  }
  return new Response("認証が必要です", {
    status: 401,
    headers: { "WWW-Authenticate": 'Basic realm="houmon"' },
  });
}
