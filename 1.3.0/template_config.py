"""Central template configuration for Allison's Data Extractor.

The live settings are stored outside the EXE and can be edited through the
administrator settings page. Constants remain for compatibility with older code,
but new code should call get_template_markers()/get_template_settings().
"""

from settings_manager import get_template_markers, get_template_settings

TEMPLATE_MARKERS = get_template_markers()
TEMPLATE_SETTINGS = get_template_settings()
