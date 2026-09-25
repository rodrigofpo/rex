"""Testes do dimensionamento responsivo da interface Desktop."""

import pytest

from rex.ui import desktop
from rex.ui.desktop import calculate_window_geometry


def test_linux_scaling_is_neutralized_before_window_creation(monkeypatch):
    calls = []
    monkeypatch.setattr(desktop.platform, "system", lambda: "Linux")
    monkeypatch.setattr(desktop.ctk, "set_widget_scaling", lambda scale: calls.append(("widget", scale)))
    monkeypatch.setattr(desktop.ctk, "set_window_scaling", lambda scale: calls.append(("window", scale)))

    def stop_at_window_creation(_self):
        assert calls == [("widget", 1.0), ("window", 1.0)]
        raise RuntimeError("janela interceptada")

    monkeypatch.setattr(desktop.ctk.CTk, "__init__", stop_at_window_creation)
    with pytest.raises(RuntimeError, match="janela interceptada"):
        desktop.REXDesktopApp()


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
