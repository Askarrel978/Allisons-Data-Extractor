import random

import pytest
from PIL import Image, ImageDraw

from pdf_extractor import _detect_grid_lines, _group_nearby


def _reference_detect(image, bright_threshold=100, horizontal_coverage=0.75, vertical_coverage=0.70):
    """The original pixel-by-pixel implementation, kept to prove identical output."""
    gray = image.convert("L")
    width, height = gray.size
    pix = gray.load()
    horizontal = [
        y for y in range(height)
        if sum(1 for x in range(width) if pix[x, y] >= bright_threshold) >= int(width * horizontal_coverage)
    ]
    vertical = [
        x for x in range(width)
        if sum(1 for y in range(height) if pix[x, y] >= bright_threshold) >= int(height * vertical_coverage)
    ]
    return _group_nearby(horizontal), _group_nearby(vertical)


def _template_b_like_grid():
    image = Image.new("RGB", (420, 360), (30, 30, 30))
    draw = ImageDraw.Draw(image)
    for y in range(10, 350, 40):
        draw.line([(0, y), (419, y)], fill=(200, 200, 200), width=2)
    for x in (2, 150, 417):
        draw.line([(x, 0), (x, 359)], fill=(200, 200, 200), width=1)
    draw.text((20, 20), "Plant Species", fill=(255, 255, 255))
    return image


def test_detects_table_rules():
    horizontal, vertical = _detect_grid_lines(_template_b_like_grid())
    assert len(horizontal) == 9
    assert vertical == [2, 150, 417]


def test_matches_original_implementation_on_grid():
    image = _template_b_like_grid()
    assert _detect_grid_lines(image) == _reference_detect(image)


@pytest.mark.parametrize("seed", [1, 2, 3])
def test_matches_original_implementation_on_noise(seed):
    # Dense noise straddling the threshold exercises the coverage boundaries.
    rng = random.Random(seed)
    width, height = 97, 61
    image = Image.new("L", (width, height))
    image.putdata([rng.choice((99, 100, 101, 255)) if rng.random() < 0.8 else 0 for _ in range(width * height)])
    assert _detect_grid_lines(image) == _reference_detect(image)
