import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://tdoubfwlpfcyxqfqyppb.supabase.co'

const supabaseKey = 'sb_publishable_0j_LkTn-MVHW88FTMFjzJA_1tX8jquO'

export const supabase = createClient(
  supabaseUrl,
  supabaseKey
)
