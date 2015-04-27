import miner_globals
# define targets
miner_globals.addExtensionToTargetMapping(".xlsx", "excel")

# always generate excel via proxy
miner_globals.addTargetToClassMapping("excel", None, "excel_target.oExcel", "Creates Excel spreadsheet")
