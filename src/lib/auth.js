import { supabase, supabaseConfigured } from "./supabase.js";

function requireSupabase() {
  if (!supabaseConfigured || !supabase) {
    throw new Error("Skolar authentication is not configured. Please contact the school administrator.");
  }
  return supabase;
}

export async function getAuthSession() {
  const client = requireSupabase();
  const { data, error } = await client.auth.getSession();
  if (error) throw error;
  return data.session;
}

export async function signInWithEmail(email, password) {
  const client = requireSupabase();
  const { data, error } = await client.auth.signInWithPassword({
    email: email.trim().toLowerCase(),
    password,
  });
  if (error) throw error;
  return data;
}

export async function registerSchoolOwner({ email, password, fullName, phone, schoolName }) {
  const client = requireSupabase();
  const { data, error } = await client.auth.signUp({
    email: email.trim().toLowerCase(),
    password,
    options: {
      data: {
        full_name: fullName.trim(),
        phone: phone.trim(),
        school_name: schoolName.trim(),
        account_intent: "create_school",
      },
    },
  });
  if (error) throw error;
  return data;
}

export async function requestPasswordReset(email, redirectTo) {
  const client = requireSupabase();
  const { error } = await client.auth.resetPasswordForEmail(
    email.trim().toLowerCase(),
    redirectTo ? { redirectTo } : undefined
  );
  if (error) throw error;
}

export async function signOutFromSkolar() {
  const client = requireSupabase();
  const { error } = await client.auth.signOut();
  if (error) throw error;
}

export function onAuthStateChange(callback) {
  const client = requireSupabase();
  return client.auth.onAuthStateChange((event, session) => callback(event, session));
}
