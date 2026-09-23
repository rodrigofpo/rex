"""Testes do dimensionamento responsivo da interface Desktop."""

import pytest

from rex.ui.desktop import calculate_linux_ui_scale, calculate_window_geometry


def test_converts_hidpi_tk_scaling_for_customtkinter():
    assert calculate_linux_ui_scale(2.669293924466338) == pytest.approx(2.002, rel=1e-3)


def test_linux_ui_scale_has_safe_limits():
    assert calculate_linux_ui_scale(1.0) == 1.0
    assert calculate_linux_ui_scale(4.0) == 2.5


def test_uses_wide_default_geometry_on_full_hd_screen():
    assert calculate_window_geometry(1920, 1080) == (1200, 820, 360, 130)


def test_caps_geometry_to_available_notebook_screen():
    width, height, offset_x, offset_y = calculate_window_geometry(1366, 768)

    assert (width, height) == (1200, 675)
    assert offset_x >= 0
    assert offset_y >= 0


def test_keeps_window_visible_on_small_screen():
    width, height, offset_x, offset_y = calculate_window_geometry(800, 600)

    assert (width, height) == (720, 600)
    assert (offset_x, offset_y) == (40, 0)
