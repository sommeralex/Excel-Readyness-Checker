"""
Prüft, ob alle Regelklassen ein korrektes rule_name-Attribut haben und dieses kein None/leer/undefined ist.
"""
from excel_checker.rules import ALL_RULES

errors = []
for rule_cls in ALL_RULES:
    try:
        rule = rule_cls()
        name = getattr(rule, 'rule_name', None)
        if not isinstance(name, str) or not name or name.lower() == 'undefined':
            errors.append(f"❌ {rule_cls.__name__}: rule_name ist ungültig: {repr(name)}")
    except Exception as e:
        errors.append(f"❌ {rule_cls.__name__}: Fehler beim Instanziieren: {e}")

if errors:
    print("Fehlerhafte Regeln:")
    for err in errors:
        print(err)
else:
    print("✅ Alle Regelklassen liefern ein gültiges rule_name-Attribut.")
