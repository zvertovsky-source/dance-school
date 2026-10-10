// Переименуйте в config.js и подставьте свои значения (Supabase → Settings → API).
// Ключ anon публичный, его можно хранить в репозитории. Ключ service_role сюда нельзя.
window.sb = supabase.createClient("https://ВАШ_PROJECT_REF.supabase.co", "ВАШ_ANON_KEY");
