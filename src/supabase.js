import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://hnjawylcwfhxjhkvtyyc.supabase.co";
const supabaseAnonKey = "sb_publishable_rBDOMqscRjKTUPvGBlevnA_P3KtnzEi";

export const supabase = createClient(supabaseUrl, supabaseAnonKey);