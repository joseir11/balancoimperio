import { createClient } from "https://esm.sh/@supabase/supabase-js@2.57.0";
const cors = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { ...cors, "Content-Type": "application/json" } });
  try {
    const authHeader = req.headers.get("Authorization");
    if (!authHeader) return json({ error: "Não autenticado." }, 401);
    const url = Deno.env.get("SUPABASE_URL")!;
    const anon = Deno.env.get("SUPABASE_ANON_KEY") ?? Deno.env.get("SUPABASE_PUBLISHABLE_KEY")!;
    const service = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    if (!service) return json({ error: "A função precisa do secret SUPABASE_SERVICE_ROLE_KEY." }, 500);
    const caller = createClient(url, anon, { global: { headers: { Authorization: authHeader } } });
    const { data: { user }, error: authError } = await caller.auth.getUser();
    if (authError || !user) return json({ error: "Sessão inválida." }, 401);
    const { data: profile } = await caller.from("profiles").select("role").eq("id", user.id).single();
    if (profile?.role !== "admin") return json({ error: "Acesso restrito ao administrador." }, 403);
    const admin = createClient(url, service, { auth: { autoRefreshToken: false, persistSession: false } });
    const body = await req.json();
    if (body.action === "create") {
      const name = String(body.name ?? "").trim(), password = String(body.password ?? "");
      if (!name || password.length < 6) return json({ error: "Informe o nome e uma senha com pelo menos 6 caracteres." }, 400);
      const slug = name.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, ".").replace(/^\.|\.$/g, "") || "equipe";
      const suffix = String(crypto.getRandomValues(new Uint16Array(1))[0] % 9000 + 1000);
      const login_code = `${slug.replace(/\./g, "-").toUpperCase()}-${suffix}`;
      const email = `${slug}.${suffix}@teams.balancoimperio.app`;
      const { data: created, error: ce } = await admin.auth.admin.createUser({ email, password, email_confirm: true, user_metadata: { display_name: name, account_type: "team" } });
      if (ce || !created.user) return json({ error: ce?.message ?? "Falha ao criar usuário." }, 400);
      const { data: team, error: te } = await admin.from("teams").insert({ name, auth_user_id: created.user.id, login_email: email, login_code }).select("id,name,login_code").single();
      if (te) { await admin.auth.admin.deleteUser(created.user.id); return json({ error: te.message }, 400); }
      const { error: pe } = await admin.from("profiles").update({ role: "team", team_id: team.id }).eq("id", created.user.id);
      if (pe) return json({ error: pe.message }, 400);
      return json({ team });
    }
    if (body.action === "update") {
      const teamId = String(body.team_id ?? ""), name = String(body.name ?? "").trim(), password = String(body.password ?? "");
      const { data: team, error } = await admin.from("teams").select("id,auth_user_id").eq("id", teamId).single();
      if (error || !team) return json({ error: "Equipe não encontrada." }, 404);
      if (name) { const { error: e } = await admin.from("teams").update({ name }).eq("id", teamId); if (e) return json({ error: e.message }, 400); }
      if (password) { if (password.length < 6) return json({ error: "A senha deve ter pelo menos 6 caracteres." }, 400); const { error: e } = await admin.auth.admin.updateUserById(team.auth_user_id, { password }); if (e) return json({ error: e.message }, 400); }
      return json({ success: true });
    }
    if (body.action === "delete") {
      const { data: team, error } = await admin.from("teams").select("id,auth_user_id").eq("id", String(body.team_id ?? "")).single();
      if (error || !team) return json({ error: "Equipe não encontrada." }, 404);
      const { error: e } = await admin.auth.admin.deleteUser(team.auth_user_id);
      if (e) return json({ error: e.message }, 400);
      return json({ success: true });
    }
    return json({ error: "Ação inválida." }, 400);
  } catch (e) { return json({ error: e instanceof Error ? e.message : "Erro interno." }, 500); }
});