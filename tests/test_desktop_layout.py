"""Testes da interface desktop e de seu dimensionamento responsivo."""

from types import SimpleNamespace

import pandas as pd
import pytest

from rex.ui import desktop
from rex.ui.desktop import calculate_window_geometry


def test_point_count_keeps_repeated_numbers_across_samples_and_series():
    data = pd.DataFrame(
        [
            ("A01", "image1.png", 1, 1, "O"),
            ("A01", "image1.png", 1, 1, "Si"),
            ("A01", "image1.png", 2, 1, "O"),
            ("A02", "image2.png", 1, 1, "O"),
        ],
        columns=["Amostra", "Micrografia", "Serie", "Ponto", "Elemento"],
    )

    assert desktop.count_eds_points(data) == 3


def test_linux_scaling_matches_tk_before_building_widgets(monkeypatch):
    calls = []
    monkeypatch.setattr(desktop.platform, "system", lambda: "Linux")
    monkeypatch.setattr(desktop.ctk, "set_widget_scaling", lambda scale: calls.append(("widget", scale)))
    monkeypatch.setattr(desktop.ctk, "set_window_scaling", lambda scale: calls.append(("window", scale)))

    class FakeTk:
        @staticmethod
        def call(*args):
            assert args == ("tk", "scaling")
            return 8 / 3

    monkeypatch.setattr(desktop.ctk.CTk, "__init__", lambda self: setattr(self, "tk", FakeTk()))
    monkeypatch.setattr(desktop.ctk.CTk, "_get_window_scaling", lambda self: 2.0)

    def stop_at_title(_self, _title):
        assert calls == [("widget", 2.0), ("window", 2.0)]
        raise RuntimeError("janela interceptada")

    monkeypatch.setattr(desktop.ctk.CTk, "title", stop_at_title)
    with pytest.raises(RuntimeError, match="janela interceptada"):
        desktop.REXDesktopApp()


@pytest.mark.parametrize("tk_scaling, expected", [(4 / 3, 1.0), (2.0, 1.5), (8 / 3, 2.0)])
def test_linux_scale_conversion(tk_scaling, expected):
    assert desktop.linux_scale_from_tk(tk_scaling) == pytest.approx(expected)


def test_scaled_geometry_uses_logical_size_and_physical_position():
    assert calculate_window_geometry(1920, 1080, 2.0) == (864, 540, 96, 0)


@pytest.mark.parametrize("delta, expected", [(-120, 1), (120, -1), (-240, 2)])
def test_linux_mousewheel_scrolls_page(delta, expected):
    calls = []
    canvas = SimpleNamespace(
        yview=lambda: (0.0, 0.5),
        yview_scroll=lambda amount, units: calls.append((amount, units)),
    )
    shell = SimpleNamespace(
        _parent_canvas=canvas,
        _check_if_valid_scroll=lambda widget: True,
    )
    app = SimpleNamespace(scroll_shell=shell)

    desktop.REXDesktopApp._scroll_linux_mousewheel(
        app, SimpleNamespace(delta=delta, widget=object())
    )

    assert calls == [(expected, "units")]


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
