"""Explicit expression subset, validated before dependency extraction."""

import re

from openpyxl.formula.tokenizer import Tokenizer, TokenizerError
from openpyxl.utils.cell import coordinate_to_tuple

FUNCTIONS = frozenset("SUM AVERAGE MIN MAX COUNT COUNTA IF IFS AND OR NOT IFERROR VLOOKUP HLOOKUP XLOOKUP ROUND ROUNDUP ROUNDDOWN ABS INT MOD TRUE FALSE NOW TODAY RAND RANDBETWEEN OFFSET INDIRECT CELL INFO".split())
VOLATILE = frozenset("NOW TODAY RAND RANDBETWEEN OFFSET INDIRECT CELL INFO".split())
CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")


def tokenize(formula):
    """Validate operands, calls, grouping and operators with a Pratt parser."""
    try:
        tokens = [t for t in Tokenizer(formula).items if t.type != "WHITE-SPACE"]
    except (TokenizerError, ValueError, IndexError) as exc:
        return {"tokens": [], "references": [], "functions": [], "constants": [],
                "operators": [], "issues": [f"Tokenization failed: {exc}"]}
    issues = []
    position = 0
    precedence = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
                  "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}

    def expression(level=0):
        nonlocal position
        if position >= len(tokens):
            raise ValueError("Missing operand")
        token = tokens[position]
        position += 1
        if token.type == "OP_PRE" and token.value in ("+", "-"):
            expression(6)
        elif token.type == "OPERAND" and token.subtype in ("RANGE", "NUMBER", "TEXT", "LOGICAL", "ERROR"):
            pass
        elif token.type in ("FUNC", "PAREN") and token.subtype == "OPEN":
            closing_type = token.type
            if position < len(tokens) and tokens[position].subtype == "CLOSE":
                if closing_type == "PAREN":
                    raise ValueError("Empty parentheses")
            else:
                expression()
                while position < len(tokens) and tokens[position].type == "SEP":
                    if closing_type != "FUNC" or tokens[position].value != ",":
                        raise ValueError("Unsupported separator")
                    position += 1
                    expression()
            if position >= len(tokens) or tokens[position].type != closing_type or tokens[position].subtype != "CLOSE":
                raise ValueError("Unbalanced call or parentheses")
            position += 1
        else:
            raise ValueError(f"Unsupported expression token {token.value!r}")
        while position < len(tokens):
            token = tokens[position]
            if token.type == "OP_POST" and token.value == "%":
                position += 1
                continue
            if token.type != "OP_IN":
                break
            if token.value not in precedence:
                raise ValueError(f"Unsupported operator {token.value}")
            rank = precedence[token.value]
            if rank < level:
                break
            position += 1
            expression(rank if token.value == "^" else rank + 1)

    try:
        expression()
        if position != len(tokens):
            raise ValueError("Trailing tokens or unsupported intersection/union")
    except (ValueError, RecursionError) as exc:
        issues.append(str(exc))
    functions = sorted({t.value[:-1].upper() for t in tokens if t.type == "FUNC" and t.subtype == "OPEN"})
    issues.extend(f"Unsupported function: {name}" for name in functions if name not in FUNCTIONS)
    issues.extend(f"Dynamic reference requires manual review: {name}" for name in functions if name in {"INDIRECT", "OFFSET"})
    return {"tokens": [{"value": t.value, "type": t.type, "subtype": t.subtype} for t in tokens],
            "references": [t.value for t in tokens if t.type == "OPERAND" and t.subtype == "RANGE"],
            "functions": functions,
            "constants": [t.value for t in tokens if t.type == "OPERAND" and t.subtype in {"NUMBER", "TEXT", "LOGICAL"}],
            "operators": [t.value for t in tokens if t.type.startswith("OP_")], "issues": issues}


def pattern(formula, coordinate):
    """Relative references preserve offsets, absolute anchors and sheet identity."""
    row, col = coordinate_to_tuple(coordinate)
    parts = []
    for token in tokenize(formula)["tokens"]:
        value = token["value"]
        if token["subtype"] == "RANGE":
            def replace(match):
                from openpyxl.utils.cell import column_index_from_string
                ca, letters, ra, number = match.groups()
                c = column_index_from_string(letters)
                r = int(number)
                return f"C{'A' + str(c) if ca else 'R' + str(c - col)}R{'A' + str(r) if ra else 'R' + str(r - row)}"
            value = re.sub(r"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_])", replace, value)
        parts.append(value)
    return "".join(parts)
