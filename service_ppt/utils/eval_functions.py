"""Evaluation functions for slide expressions.

This module provides functions and classes for evaluating expressions that
match or select slides in a presentation, supporting safe expression evaluation
with slide context.
"""

import typing

from service_ppt.utils.evaluator import safe_eval


class EvalShape:
    """Support slide matching logic with contains_text method.

    This class provides a way to check if text exists in a slide or its notes,
    supporting the slide matching logic used in expression evaluation.
    """

    def __init__(self, prs, slide_index, note_shapes):
        """Initialize EvalShape.

        :param prs: Presentation object
        :param slide_index: Index of the slide to evaluate
        :param note_shapes: If True, search in notes; if False, search in slide
        """
        self.prs = prs
        self.slide_index = slide_index
        self.note_shapes = note_shapes

    def contains_text(self, text, ignore_case=False, whole_words=False):
        """Check if the slide or note contains the specified text.

        :param text: Text to search for
        :param ignore_case: If True, perform case-insensitive search
        :param whole_words: If True, match whole words only
        :returns: True if text is found, False otherwise
        """
        return self.prs.find_text_in_slide(self.slide_index, self.note_shapes, text, ignore_case, whole_words)


def populate_slide_dict(prs, slide_index):
    """Construct dictionary for slide expression evaluation.

    Creates a dictionary with 'slide' and 'note' keys that can be used in
    expression evaluation to match slides based on their content.

    :param prs: Presentation object
    :param slide_index: Index of the slide to create dictionary for
    :returns: Dictionary with 'slide' and 'note' EvalShape instances
    """
    sdict = {
        "slide": EvalShape(prs, slide_index, False),
        "note": EvalShape(prs, slide_index, True),
    }

    return sdict


def evaluate_to_single_slide(prs, expr) -> typing.Any | None:
    """Evaluate expression to find a single matching slide.

    Evaluates an expression that should match a single slide. If the expression
    doesn't depend on slide context, returns the result directly. Otherwise,
    evaluates the expression for each slide and returns the index of the first
    matching slide.

    :param prs: Presentation object
    :param expr: Expression string to evaluate
    :returns: Slide index if match found, expression result if no slide dependency, or None
    """
    if not expr:
        return None

    # if the expr works with empty dict meaning it doesn't have dependency to slide dict,
    # return it.
    try:
        gdict = {}
        result = safe_eval(expr, gdict)
        return result
    except (NameError, ValueError, SyntaxError, AttributeError, TypeError):
        # Expression depends on slide context, continue to slide-by-slide evaluation
        # This is expected when expression requires slide-specific variables
        pass

    # Use each slide's dict to evaluate the expr.
    for index in range(prs.slide_count()):
        gdict = populate_slide_dict(prs, index)

        try:
            eval_result = safe_eval(expr, gdict)
            if eval_result:
                return index
        except (NameError, ValueError, SyntaxError, AttributeError, TypeError):
            # Skip this slide if evaluation fails
            # This is expected when expression cannot be evaluated for this slide
            pass

    return None


def evaluate_to_multiple_slide(prs, expr):
    """Evaluate expression to find multiple matching slides.

    Evaluates an expression that should match multiple slides. If the expression
    doesn't depend on slide context, returns the result directly (if it's a list).
    Otherwise, evaluates the expression for each slide and returns a list of
    indices of matching slides.

    :param prs: Presentation object
    :param expr: Expression string to evaluate
    :returns: List of slide indices if matches found, expression result if no slide dependency, or None
    """
    if expr is None:
        return None

    # if the expr works with empty dict meaning it doesn't have dependency to slide dict,
    # return it.
    try:
        gdict = {}
        result = safe_eval(expr, gdict)
        # If result is an empty list, return None to match original behavior
        if isinstance(result, list) and len(result) == 0:
            return None
        return result
    except (NameError, ValueError, SyntaxError, AttributeError, TypeError):
        # Expression depends on slide context, continue to slide-by-slide evaluation
        # This is expected when expression requires slide-specific variables
        pass

    # Use each slide's dict to evaluate the expr.
    result = []
    for index in range(prs.slide_count()):
        gdict = populate_slide_dict(prs, index)

        try:
            eval_result = safe_eval(expr, gdict)
            if eval_result:
                result.append(index)
        except (NameError, ValueError, SyntaxError, AttributeError, TypeError):
            # Skip this slide if evaluation fails
            # This is expected when expression cannot be evaluated for this slide
            pass

    if len(result) == 0:
        return None

    return result
