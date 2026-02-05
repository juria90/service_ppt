"""Tests for evaluator module.

This module contains unit tests for the safe expression evaluator.
"""

import pytest

from service_ppt.utils.evaluator import SafeEvaluator, safe_eval


class TestSafeEvaluator:
    """Test SafeEvaluator class."""

    def test_evaluates_constants(self):
        """Test that SafeEvaluator evaluates constants correctly."""
        evaluator = SafeEvaluator({})

        assert evaluator.evaluate("42") == 42
        assert evaluator.evaluate("3.14") == 3.14
        assert evaluator.evaluate("'hello'") == "hello"
        assert evaluator.evaluate("True") is True
        assert evaluator.evaluate("False") is False
        assert evaluator.evaluate("None") is None

    def test_evaluates_arithmetic(self):
        """Test that SafeEvaluator evaluates arithmetic expressions."""
        evaluator = SafeEvaluator({})

        assert evaluator.evaluate("2 + 3") == 5
        assert evaluator.evaluate("10 - 4") == 6
        assert evaluator.evaluate("3 * 4") == 12
        assert evaluator.evaluate("15 / 3") == 5.0
        assert evaluator.evaluate("17 // 5") == 3
        assert evaluator.evaluate("17 % 5") == 2
        assert evaluator.evaluate("2 ** 3") == 8

    def test_evaluates_comparisons(self):
        """Test that SafeEvaluator evaluates comparison expressions."""
        evaluator = SafeEvaluator({})

        assert evaluator.evaluate("5 > 3") is True
        assert evaluator.evaluate("2 < 1") is False
        assert evaluator.evaluate("10 == 10") is True
        assert evaluator.evaluate("5 != 3") is True
        assert evaluator.evaluate("4 >= 4") is True
        assert evaluator.evaluate("3 <= 2") is False

    def test_evaluates_boolean_operations(self):
        """Test that SafeEvaluator evaluates boolean operations."""
        evaluator = SafeEvaluator({})

        assert evaluator.evaluate("True and True") is True
        assert evaluator.evaluate("True and False") is False
        assert evaluator.evaluate("True or False") is True
        assert evaluator.evaluate("False or False") is False
        assert evaluator.evaluate("not False") is True

    def test_evaluates_lists(self):
        """Test that SafeEvaluator evaluates list expressions."""
        evaluator = SafeEvaluator({})

        assert evaluator.evaluate("[1, 2, 3]") == [1, 2, 3]
        assert evaluator.evaluate("[]") == []
        assert evaluator.evaluate("[1, 'a', True]") == [1, "a", True]

    def test_evaluates_dicts(self):
        """Test that SafeEvaluator evaluates dictionary expressions."""
        evaluator = SafeEvaluator({})

        assert evaluator.evaluate("{'a': 1, 'b': 2}") == {"a": 1, "b": 2}
        assert evaluator.evaluate("{}") == {}

    def test_evaluates_with_context(self):
        """Test that SafeEvaluator uses context variables."""
        context = {"x": 10, "y": 20}
        evaluator = SafeEvaluator(context)

        assert evaluator.evaluate("x + y") == 30
        assert evaluator.evaluate("x * 2") == 20

    def test_raises_on_undefined_variable(self):
        """Test that SafeEvaluator raises NameError for undefined variables."""
        evaluator = SafeEvaluator({})

        with pytest.raises(NameError, match="Name 'x' is not defined"):
            evaluator.evaluate("x + 1")

    def test_raises_on_empty_expression(self):
        """Test that SafeEvaluator raises SyntaxError for empty expression."""
        evaluator = SafeEvaluator({})

        with pytest.raises(SyntaxError, match="Empty expression"):
            evaluator.evaluate("")

        with pytest.raises(SyntaxError, match="Empty expression"):
            evaluator.evaluate("   ")

    def test_raises_on_invalid_syntax(self):
        """Test that SafeEvaluator raises SyntaxError for invalid syntax."""
        evaluator = SafeEvaluator({})

        with pytest.raises(SyntaxError):
            evaluator.evaluate("2 +")

    def test_evaluates_attribute_access(self):
        """Test that SafeEvaluator evaluates attribute access."""
        class TestObj:
            def __init__(self):
                self.value = 42

        context = {"obj": TestObj()}
        evaluator = SafeEvaluator(context)

        assert evaluator.evaluate("obj.value") == 42

    def test_evaluates_function_calls(self):
        """Test that SafeEvaluator evaluates function calls."""
        def add(a, b):
            return a + b

        context = {"add": add}
        evaluator = SafeEvaluator(context)

        assert evaluator.evaluate("add(2, 3)") == 5

    def test_evaluates_subscripts(self):
        """Test that SafeEvaluator evaluates subscript operations."""
        context = {"lst": [1, 2, 3], "d": {"a": 1}}
        evaluator = SafeEvaluator(context)

        assert evaluator.evaluate("lst[0]") == 1
        assert evaluator.evaluate("d['a']") == 1

    def test_evaluates_conditional_expressions(self):
        """Test that SafeEvaluator evaluates conditional expressions."""
        evaluator = SafeEvaluator({})

        assert evaluator.evaluate("5 if True else 10") == 5
        assert evaluator.evaluate("5 if False else 10") == 10


class TestSafeEval:
    """Test safe_eval convenience function."""

    def test_safe_eval_works_like_evaluator(self):
        """Test that safe_eval works like SafeEvaluator.evaluate."""
        context = {"x": 10}

        result = safe_eval("x + 5", context)
        assert result == 15

    def test_safe_eval_raises_on_error(self):
        """Test that safe_eval raises errors like SafeEvaluator."""
        with pytest.raises(NameError):
            safe_eval("undefined_var", {})
