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
        return {"tokens": [], "calls": [], "references": [], "functions": [], "constants": [],
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
        if token.type == "OPERATOR-PREFIX" and token.value in ("+", "-"):
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
            if token.type == "OPERATOR-POSTFIX" and token.value == "%":
                position += 1
                continue
            if token.type != "OPERATOR-INFIX":
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
    calls = []
    for offset, token in enumerate(tokens):
        if token.type != "FUNC" or token.subtype != "OPEN":
            continue
        args, current, depth = [], [], 0
        for other in tokens[offset + 1:]:
            if other.subtype == "CLOSE" and depth == 0:
                if current:
                    args.append("".join(current))
                break
            if other.type == "SEP" and depth == 0:
                args.append("".join(current))
                current = []
                continue
            if other.subtype == "OPEN":
                depth += 1
            elif other.subtype == "CLOSE":
                depth -= 1
            current.append(other.value)
        calls.append({"function": token.value[:-1].upper(), "arguments": args})
    return {"tokens": [{"value": t.value, "type": t.type, "subtype": t.subtype} for t in tokens], "calls": calls,
            "references": [t.value for t in tokens if t.type == "OPERAND" and t.subtype == "RANGE"],
            "functions": functions,
            "constants": [t.value for t in tokens if t.type == "OPERAND" and t.subtype in {"NUMBER", "TEXT", "LOGICAL"}],
            "operators": [t.value for t in tokens if t.type.startswith("OPERATOR-")], "issues": issues}


def pattern(formula, coordinate):
    """Relative references preserve offsets, absolute anchors and sheet identity."""
    row, col = coordinate_to_tuple(coordinate)
    parts = []
    for token in tokenize(formula)["tokens"]:
        value = token["value"]
        if token["subtype"] == "RANGE":
            prefix, separator, address = value.rpartition("!")
            if not separator:
                address = value
            def replace(match):
                from openpyxl.utils.cell import column_index_from_string
                ca, letters, ra, number = match.groups()
                c = column_index_from_string(letters)
                r = int(number)
                return f"C{'A' + str(c) if ca else 'R' + str(c - col)}R{'A' + str(r) if ra else 'R' + str(r - row)}"
            address = re.sub(r"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_])", replace, address)
            value = prefix + separator + address if separator else address
        parts.append(value)
    return "".join(parts)
