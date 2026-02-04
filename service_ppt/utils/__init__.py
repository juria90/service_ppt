"""Utility modules for service_ppt.

This package contains utility functions and classes used across the service_ppt application.
"""

from service_ppt.utils.atomicfile import AtomicFileWriter
from service_ppt.utils.eval_functions import (
    EvalShape,
    evaluate_to_multiple_slide,
    evaluate_to_single_slide,
    populate_slide_dict,
)
from service_ppt.utils.i18n import _, get_translation, initialize_translation, ngettext, pgettext
from service_ppt.utils.make_transparent import color_to_transparent
from service_ppt.utils.process_exists import process_exists
from service_ppt.utils.evaluator import safe_eval

__all__ = [
    "_",
    "AtomicFileWriter",
    "color_to_transparent",
    "EvalShape",
    "evaluate_to_multiple_slide",
    "evaluate_to_single_slide",
    "get_translation",
    "initialize_translation",
    "ngettext",
    "populate_slide_dict",
    "process_exists",
    "pgettext",
    "safe_eval",
]
