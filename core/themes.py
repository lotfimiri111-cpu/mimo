"""
Theme System — مذكرتي Pro v17
All color palettes defined in one place. Immutable after import.
"""
from __future__ import annotations
from dataclasses import dataclass
from pptx.dml.color import RGBColor


def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


@dataclass(frozen=True)
class Theme:
    name: str
    family: str

    # Main palette
    bg: str           # Background (dark)
    bg2: str          # Background secondary
    accent: str       # Primary accent (e.g. gold)
    accent2: str      # Secondary accent
    text_light: str   # Text on dark bg
    text_dark: str    # Text on light bg
    card: str         # Card background
    muted: str        # Muted/subtle color

    # Gradient stops
    grad1: str
    grad2: str
    accent_grad1: str
    accent_grad2: str

    def rgb(self, attr: str) -> RGBColor:
        return _rgb(getattr(self, attr))

    @property
    def bg_rgb(self): return _rgb(self.bg)
    @property
    def bg2_rgb(self): return _rgb(self.bg2)
    @property
    def accent_rgb(self): return _rgb(self.accent)
    @property
    def accent2_rgb(self): return _rgb(self.accent2)
    @property
    def text_light_rgb(self): return _rgb(self.text_light)
    @property
    def text_dark_rgb(self): return _rgb(self.text_dark)
    @property
    def card_rgb(self): return _rgb(self.card)
    @property
    def muted_rgb(self): return _rgb(self.muted)


THEMES: dict[str, Theme] = {
    "navy_gold": Theme(
        name="navy_gold", family="Navy & Gold",
        bg="#07172F", bg2="#0D2347",
        accent="#C6A03C", accent2="#E8C96A",
        text_light="#F0E8D8", text_dark="#07172F",
        card="#122040", muted="#8A9BB5",
        grad1="#07172F", grad2="#1A3A6E",
        accent_grad1="#C6A03C", accent_grad2="#E8C96A",
    ),
    "dark_teal": Theme(
        name="dark_teal", family="Dark Teal",
        bg="#0A1F1C", bg2="#0F2E29",
        accent="#2EBFA5", accent2="#5DD4BE",
        text_light="#E8F5F3", text_dark="#0A1F1C",
        card="#122820", muted="#7BAAA4",
        grad1="#0A1F1C", grad2="#1A4A42",
        accent_grad1="#2EBFA5", accent_grad2="#5DD4BE",
    ),
    "burgundy": Theme(
        name="burgundy", family="Burgundy",
        bg="#1A0A0F", bg2="#2D1018",
        accent="#C0392B", accent2="#E74C3C",
        text_light="#F5E8E8", text_dark="#1A0A0F",
        card="#2A1015", muted="#A07878",
        grad1="#1A0A0F", grad2="#3D1520",
        accent_grad1="#C0392B", accent_grad2="#E74C3C",
    ),
    "forest": Theme(
        name="forest", family="Forest",
        bg="#0D1F0D", bg2="#152815",
        accent="#27AE60", accent2="#58D68D",
        text_light="#E8F5E8", text_dark="#0D1F0D",
        card="#152215", muted="#7AAE7A",
        grad1="#0D1F0D", grad2="#1E4020",
        accent_grad1="#27AE60", accent_grad2="#58D68D",
    ),
    "midnight_purple": Theme(
        name="midnight_purple", family="Midnight Purple",
        bg="#0F0A1F", bg2="#1A1030",
        accent="#8E44AD", accent2="#BB8FCE",
        text_light="#F0E8FF", text_dark="#0F0A1F",
        card="#1A1028", muted="#8A7AAA",
        grad1="#0F0A1F", grad2="#2C1A50",
        accent_grad1="#8E44AD", accent_grad2="#BB8FCE",
    ),
    "charcoal_orange": Theme(
        name="charcoal_orange", family="Charcoal Orange",
        bg="#1A1A1A", bg2="#282828",
        accent="#E67E22", accent2="#F39C12",
        text_light="#F5F0E8", text_dark="#1A1A1A",
        card="#222222", muted="#909090",
        grad1="#1A1A1A", grad2="#3A3A3A",
        accent_grad1="#E67E22", accent_grad2="#F39C12",
    ),
    "ice_blue": Theme(
        name="ice_blue", family="Ice Blue",
        bg="#0A1520", bg2="#0F2035",
        accent="#3498DB", accent2="#85C1E9",
        text_light="#E8F4FD", text_dark="#0A1520",
        card="#122030", muted="#6A9EB5",
        grad1="#0A1520", grad2="#1A3A5A",
        accent_grad1="#3498DB", accent_grad2="#85C1E9",
    ),
    "sand_gold": Theme(
        name="sand_gold", family="Sand Gold",
        bg="#1F1505", bg2="#2E2008",
        accent="#D4AC0D", accent2="#F1C40F",
        text_light="#FDF5E6", text_dark="#1F1505",
        card="#28180A", muted="#A0906A",
        grad1="#1F1505", grad2="#3D2A08",
        accent_grad1="#D4AC0D", accent_grad2="#F1C40F",
    ),
    "slate_crimson": Theme(
        name="slate_crimson", family="Slate Crimson",
        bg="#1A1A2E", bg2="#16213E",
        accent="#E94560", accent2="#F07070",
        text_light="#F0F0F8", text_dark="#1A1A2E",
        card="#1E2040", muted="#8A8AB0",
        grad1="#1A1A2E", grad2="#0F3460",
        accent_grad1="#E94560", accent_grad2="#F07070",
    ),
    "noir": Theme(
        name="noir", family="Noir",
        bg="#0D0D0D", bg2="#1A1A1A",
        accent="#C9B99A", accent2="#E8D5B5",
        text_light="#F5F5F5", text_dark="#0D0D0D",
        card="#1A1A1A", muted="#888888",
        grad1="#0D0D0D", grad2="#2A2A2A",
        accent_grad1="#C9B99A", accent_grad2="#E8D5B5",
    ),
    "atlas": Theme(
        name="atlas", family="Atlas",
        bg="#0A0F1E", bg2="#101828",
        accent="#00BCD4", accent2="#4DD0E1",
        text_light="#E8F8FF", text_dark="#0A0F1E",
        card="#121F35", muted="#6A9BAB",
        grad1="#0A0F1E", grad2="#1A2F4E",
        accent_grad1="#00BCD4", accent_grad2="#4DD0E1",
    ),
    "sakura": Theme(
        name="sakura", family="Sakura",
        bg="#1A0A14", bg2="#2A1020",
        accent="#E91E8C", accent2="#F06EB5",
        text_light="#FFE8F5", text_dark="#1A0A14",
        card="#22102A", muted="#A07090",
        grad1="#1A0A14", grad2="#3A1530",
        accent_grad1="#E91E8C", accent_grad2="#F06EB5",
    ),
}


def get_theme(name: str) -> Theme:
    return THEMES.get(name, THEMES["navy_gold"])
