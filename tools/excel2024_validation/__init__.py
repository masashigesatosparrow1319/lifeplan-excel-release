"""Gate 2 only: tiny -> 24 formulas -> one existing annual S1 case."""

SCHEMA = 'EXCEL2024_GATE2_V1'
STAGES = ('tiny3', 'functions24', 'annual1')
E1_DIAGNOSTIC = 'E1_FORCE_FULL_FALSE'
E1_PROTOCOL = 'e1_force_false_full_rebuild_auto_wait_v1'
F24_DIAGNOSTIC = 'FUNCTIONS24_FORCE_FULL_FALSE'
F24_PROTOCOL = 'functions24_force_false_full_rebuild_auto_wait_v1'
A_DIAGNOSTIC = 'FUNCTIONS24_PREOPEN_FORCE0_A_V1'
A_PROTOCOL = 'functions24_preopen_force0_full_rebuild_auto_wait_v1'
A_TRANSFORM = 'G2_F24_PREOPEN_FORCE0_TRANSFORM_V1'
