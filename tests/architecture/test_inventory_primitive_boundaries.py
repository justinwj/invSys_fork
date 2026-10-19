"""D8-INV-01: Inventory creation crosses XLAMs using primitive envelopes.

Resolve the actual Core procedure signatures called by the affected Admin
entry points. This is architecture evidence; it supplements operator tests.
"""

from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[2]
SCALARS = {
    "string", "boolean", "byte", "integer", "long", "longlong", "longptr",
    "single", "double", "currency", "decimal", "date", "variant",
}
START = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?(Function|Sub)\s+(\w+)\b", re.I
)


def procedures(path):
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    result = {}
    for index, line in enumerate(lines):
        match = START.match(line)
        if not match:
            continue
        end = index
        signature = line
        while lines[end].rstrip().endswith("_"):
            end += 1
            signature += " " + lines[end]
        body_end = end + 1
        while body_end < len(lines):
            if re.match(r"^\s*End\s+" + match[2] + r"\s*$", lines[body_end], re.I):
                break
            body_end += 1
        if body_end == len(lines):
            raise ValueError("Incomplete VBA procedure: " + match[3])
        result[match[3].lower()] = {
            "signature": signature,
            "body": "\n".join(lines[end + 1:body_end]),
            "line": index + 1,
        }
    return result


def nonprimitive_types(signature):
    return sorted({
        name for name in re.findall(r"\bAs\s+([\w.]+)", signature, re.I)
        if name.lower() not in SCALARS
    })


class InventoryPrimitiveBoundaries(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.core = {
            path.stem.lower(): procedures(path)
            for path in (ROOT / "src/Core/Modules").glob("*.bas")
        }
        if "modroleeventwriter" not in cls.core:
            raise RuntimeError("Required local Core source is unavailable.")

    def assert_primitive_calls(self, relative_path, name):
        source = procedures(ROOT / relative_path)
        if name.lower() not in source:
            raise RuntimeError("Required Inventory entry source is unavailable: " + name)
        entry = source[name.lower()]
        violations = []
        for module, procedure in re.findall(r"\b(mod\w+)\.(\w+)\b", entry["body"], re.I):
            declarations = self.core.get(module.lower())
            if declarations is None:
                continue  # Same-project Admin calls are typed local calls.
            callee = declarations.get(procedure.lower())
            if callee is None:
                raise RuntimeError("Required Core declaration is unavailable: " + module + "." + procedure)
            types = nonprimitive_types(callee["signature"])
            if types:
                violations.append(module + "." + procedure + ": " + ", ".join(types))
        self.assertEqual([], sorted(set(violations)), name + " crosses object contracts")

    def test_add_command_uses_primitive_core_contracts(self):
        self.assert_primitive_calls(
            "src/Admin/Modules/modAdmin.bas", "AddInventoryItemForWarehouse"
        )

    def test_add_payload_builder_uses_primitive_core_contracts(self):
        self.assert_primitive_calls(
            "src/Admin/Modules/modAdmin.bas", "BuildInventoryCreatePayloadAdmin"
        )

    def test_catalog_completion_uses_primitive_core_contracts(self):
        self.assert_primitive_calls(
            "src/Admin/Modules/modAdminInventoryCreate.bas", "CompleteCatalogItem"
        )

    def test_string_key_factory_is_a_valid_primitive_control(self):
        declaration = self.core["modroleeventwriter"]["createsystemkey"]
        self.assertEqual([], nonprimitive_types(declaration["signature"]))


if __name__ == "__main__":
    unittest.main()
