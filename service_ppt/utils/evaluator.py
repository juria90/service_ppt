"""Safe expression evaluator to replace eval().

This module provides a safe way to evaluate Python expressions without using eval(),
preventing code injection vulnerabilities while supporting the expression patterns
used in the codebase.
"""

import ast
import operator
from typing import Any


class SafeEvaluator:
    """Safe expression evaluator that only allows safe operations."""

    # Allowed operations
    _OPERATORS = {
        ast.Add: operator.add,
        ast.Sub: operator.sub,
        ast.Mult: operator.mul,
        ast.Div: operator.truediv,
        ast.FloorDiv: operator.floordiv,
        ast.Mod: operator.mod,
        ast.Pow: operator.pow,
        ast.LShift: operator.lshift,
        ast.RShift: operator.rshift,
        ast.BitOr: operator.or_,
        ast.BitXor: operator.xor,
        ast.BitAnd: operator.and_,
        ast.Lt: operator.lt,
        ast.LtE: operator.le,
        ast.Gt: operator.gt,
        ast.GtE: operator.ge,
        ast.Eq: operator.eq,
        ast.NotEq: operator.ne,
        ast.Is: operator.is_,
        ast.IsNot: operator.is_not,
        ast.In: lambda a, b: a in b,
        ast.NotIn: lambda a, b: a not in b,
        ast.And: lambda a, b: a and b,
        ast.Or: lambda a, b: a or b,
        ast.Not: operator.not_,
        ast.USub: operator.neg,
        ast.UAdd: operator.pos,
        ast.Invert: operator.invert,
    }

    def __init__(self, context: dict[str, Any]):
        """Initialize evaluator with a context dictionary.

        :param context: Dictionary containing variables and objects available in expressions
        """
        self.context = context

    def evaluate(self, expr: str) -> Any:
        """Safely evaluate an expression string.

        :param expr: Expression string to evaluate
        :return: Result of the evaluation
        :raises ValueError: If expression contains unsafe operations
        :raises SyntaxError: If expression is not valid Python syntax
        """
        if not expr or not expr.strip():
            raise SyntaxError("Empty expression")

        try:
            # Parse the expression into an AST
            tree = ast.parse(expr, mode="eval")
        except SyntaxError as e:
            # Re-raise with clearer message
            raise SyntaxError(f"Invalid expression syntax: {expr}") from e

        # Evaluate the AST safely
        return self._eval_node(tree.body)

    def _eval_node(self, node: ast.AST) -> Any:
        """Recursively evaluate an AST node.

        :param node: AST node to evaluate
        :return: Result of the evaluation
        :raises ValueError: If node contains unsafe operations
        """
        # Handle constants (Python 3.8+)
        # Handle constants (Python 3.8+)
        # Since we require Python 3.12+, we only need to handle ast.Constant
        if isinstance(node, ast.Constant):
            return node.value

        if isinstance(node, ast.Name):
            if node.id in self.context:
                return self.context[node.id]
            raise NameError(f"Name '{node.id}' is not defined")

        if isinstance(node, ast.BinOp):
            left = self._eval_node(node.left)
            right = self._eval_node(node.right)
            op = self._OPERATORS.get(type(node.op))
            if op is None:
                raise ValueError(f"Unsupported binary operation: {type(node.op).__name__}")
            return op(left, right)

        if isinstance(node, ast.UnaryOp):
            operand = self._eval_node(node.operand)
            op = self._OPERATORS.get(type(node.op))
            if op is None:
                raise ValueError(f"Unsupported unary operation: {type(node.op).__name__}")
            return op(operand)

        if isinstance(node, ast.Compare):
            left = self._eval_node(node.left)
            for op, comparator in zip(node.ops, node.comparators):
                right = self._eval_node(comparator)
                op_func = self._OPERATORS.get(type(op))
                if op_func is None:
                    raise ValueError(f"Unsupported comparison operation: {type(op).__name__}")
                result = op_func(left, right)
                if not result:
                    return False
                left = right
            return True

        if isinstance(node, ast.BoolOp):
            op = self._OPERATORS.get(type(node.op))
            if op is None:
                raise ValueError(f"Unsupported boolean operation: {type(node.op).__name__}")
            values = [self._eval_node(v) for v in node.values]
            result = values[0]
            for value in values[1:]:
                result = op(result, value)
            return result

        if isinstance(node, ast.List):
            return [self._eval_node(elem) for elem in node.elts]

        if isinstance(node, ast.Tuple):
            return tuple(self._eval_node(elem) for elem in node.elts)

        if isinstance(node, ast.Dict):
            keys = [self._eval_node(k) for k in node.keys]
            values = [self._eval_node(v) for v in node.values]
            return dict(zip(keys, values))

        if isinstance(node, ast.Attribute):
            obj = self._eval_node(node.value)
            # Only allow attribute access on objects in context or their attributes
            if hasattr(obj, node.attr):
                return getattr(obj, node.attr)
            raise AttributeError(f"'{type(obj).__name__}' object has no attribute '{node.attr}'")

        if isinstance(node, ast.Call):
            func = self._eval_node(node.func)
            args = [self._eval_node(arg) for arg in node.args]
            keywords = {kw.arg: self._eval_node(kw.value) for kw in node.keywords if kw.arg is not None}

            # Only allow calls on objects from context (like slide.contains_text)
            if callable(func):
                return func(*args, **keywords)
            raise TypeError(f"'{type(func).__name__}' object is not callable")

        if isinstance(node, ast.Subscript):
            value = self._eval_node(node.value)
            # Python 3.9+ uses node.slice directly, not wrapped in ast.Index
            index = self._eval_node(node.slice)
            return value[index]

        if isinstance(node, ast.IfExp):
            test = self._eval_node(node.test)
            if test:
                return self._eval_node(node.body)
            return self._eval_node(node.orelse)

        raise ValueError(f"Unsupported AST node type: {type(node).__name__}")


def safe_eval(expr: str, context: dict[str, Any]) -> Any:
    """Safely evaluate an expression with a given context.

    This is a convenience function that creates a SafeEvaluator and evaluates the expression.

    :param expr: Expression string to evaluate
    :param context: Dictionary containing variables and objects available in expressions
    :return: Result of the evaluation
    :raises ValueError: If expression contains unsafe operations
    :raises SyntaxError: If expression is not valid Python syntax
    """
    evaluator = SafeEvaluator(context)
    return evaluator.evaluate(expr)
