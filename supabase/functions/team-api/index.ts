import { createClient } from "https://esm.sh/@supabase/supabase-js@2.57.0";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  const json = (body: unknown, status = 200) =>
    new Response(JSON.stringify(body), { status, headers: { ...cors, "Content-Type": "application/json" } });

  try {
    if (req.method !== "POST") return json({ error: "Metodo nao permitido." }, 405);

    const url = Deno.env.get("SUPABASE_URL")!;
    const anon = Deno.env.get("SUPABASE_ANON_KEY") ?? Deno.env.get("SUPABASE_PUBLISHABLE_KEY")!;
    const service = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    if (!service) return json({ error: "Funcao sem SERVICE_ROLE_KEY." }, 500);

    const admin = createClient(url, service, { auth: { autoRefreshToken: false, persistSession: false } });
    const body = await req.json().catch(() => ({} as Record<string, unknown>));
    const action = String(body.action ?? "");

    if (action === "login") {
      const name = String(body.name ?? body.team_name ?? body.team ?? "").trim();
      const code = String(body.code ?? body.login_code ?? body.access_code ?? body.accessCode ?? "").trim();
      const password = String(body.password ?? body.senha ?? "");
      const lookup = name || code;
      if (!lookup) return json({ error: "Codigo de acesso nao informado." }, 400);
      if (!password) return json({ error: "Informe a senha da equipe." }, 400);

      let team: {
        id: string;
        name: string;
        login_email: string | null;
        login_code: string;
        active: boolean;
      } | null = null;

      const byName = await admin
        .from("teams")
        .select("id,name,login_email,login_code,active")
        .ilike("name", lookup)
        .limit(1);
      team = byName.data?.[0] ?? null;

      if (!team) {
        const byCode = await admin
          .from("teams")
          .select("id,name,login_email,login_code,active")
          .ilike("login_code", lookup)
          .limit(1);
        team = byCode.data?.[0] ?? null;
      }

      if (!team || !team.active) {
        return json({ error: "Codigo de acesso invalido ou equipe desativada." }, 401);
      }
      if (!team.login_email) {
        return json({ error: "Equipe sem e-mail de login configurado." }, 400);
      }

      const authClient = createClient(url, anon, { auth: { autoRefreshToken: false, persistSession: false } });
      const { data: session, error: se } = await authClient.auth.signInWithPassword({
        email: team.login_email,
        password,
      });
      if (se || !session.session) {
        return json({ error: "Senha incorreta." }, 401);
      }

      return json({
        access_token: session.session.access_token,
        token: session.session.access_token,
        refresh_token: session.session.refresh_token,
        team: { id: team.id, name: team.name, login_code: team.login_code },
      });
    }

    const authHeader = req.headers.get("Authorization");
    if (!authHeader) return json({ error: "Nao autenticado." }, 401);

    const caller = createClient(url, anon, { global: { headers: { Authorization: authHeader } } });
    const { data: { user }, error: authError } = await caller.auth.getUser();
    if (authError || !user) return json({ error: "Sessao invalida." }, 401);

    const { data: profile } = await admin.from("profiles").select("role,team_id").eq("id", user.id).single();
    if (!profile?.team_id || profile.role !== "team") {
      return json({ error: "Acesso restrito a equipes." }, 403);
    }

    const teamId = profile.team_id as string;
    const { data: teamRow } = await admin
      .from("teams")
      .select("id,name,login_code,active")
      .eq("id", teamId)
      .single();
    if (!teamRow?.active) return json({ error: "Equipe desativada." }, 403);

    if (action === "data") {
      const [matRes, countRes, statusRes] = await Promise.all([
        admin.from("materials").select("id,code,description,unit").eq("active", true).order("code"),
        admin.from("counts").select("id,material_id,quantity,updated_at").eq("team_id", teamId),
        admin.from("item_status").select("material_id,status,teams_count"),
      ]);
      const counts = (countRes.data || []).map((c) => ({ ...c, observation: null }));
      return json({
        team: teamRow,
        materials: matRes.data || [],
        counts,
        statuses: statusRes.data || [],
      });
    }

    if (action === "save") {
      const material_id = String(body.material_id ?? "");
      const quantity = Number(body.quantity);
      if (!material_id || !Number.isFinite(quantity) || quantity < 0) {
        return json({ error: "Dados de contagem invalidos." }, 400);
      }
      const payload: Record<string, unknown> = {
        material_id,
        team_id: teamId,
        quantity,
        updated_at: new Date().toISOString(),
      };

      const result = await admin
        .from("counts")
        .upsert(payload, { onConflict: "material_id,team_id" })
        .select("id")
        .single();

      if (result.error) return json({ error: result.error.message }, 400);
      return json({ success: true, count: result.data });
    }

    if (action === "delete") {
      const count_id = String(body.count_id ?? "");
      if (!count_id) return json({ error: "Contagem nao informada." }, 400);
      const { error } = await admin.from("counts").delete().eq("id", count_id).eq("team_id", teamId);
      if (error) return json({ error: error.message }, 400);
      return json({ success: true });
    }

    return json({ error: "Acao invalida." }, 400);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : "Erro interno." }, 500);
  }
});
