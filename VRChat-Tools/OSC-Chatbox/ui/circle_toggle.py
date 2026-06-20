# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CIRCLE TOGGLE WIDGET
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Reusable smooth anti-aliased circle toggle for tkinter UIs.
# Renders at 4x resolution then downsamples via LANCZOS for crisp edges.
#
# Usage:
#   from ui.circle_toggle import CircleToggle
#
#   toggle = CircleToggle(parent, enabled=True, command=my_callback, bg="your_panel_colour")
#   toggle.pack()
#
#   # Read state
#   is_on = toggle.get()
#
#   # Set state programmatically (does NOT fire command)
#   toggle.set(False)
#
# Constructor kwargs:
#   parent   — tkinter parent widget
#   enabled  — initial boolean state (default True)
#   command  — callable(bool) fired on user click, receives new state
#   bg       — canvas background colour; match the parent panel so the
#               transparent PNG blends seamlessly (default "#1f102a")
#   color    — accent colour for the filled/outline circle (default "#a78bfa")
#   size     — diameter of the canvas in pixels (default 20)
#   pad      — inset padding inside the canvas (default 3)
#   **kwargs — passed through to tk.Canvas
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import tkinter as tk

from PIL import Image, ImageDraw, ImageTk


class CircleToggle(tk.Canvas):
    """
    Anti-aliased circle toggle widget.

    Filled circle  = enabled / ON
    Outline circle = disabled / OFF
    """

    DEFAULT_SIZE  = 20
    DEFAULT_PAD   = 3
    DEFAULT_COLOR = "#a78bfa"
    DEFAULT_BG    = "#1f102a"   # fallback; callers should pass their own PANEL colour

    def __init__(
        self,
        parent,
        *,
        enabled: bool = True,
        command=None,
        bg: str = DEFAULT_BG,
        color: str = DEFAULT_COLOR,
        size: int = DEFAULT_SIZE,
        pad: int = DEFAULT_PAD,
        **kwargs,
    ):
        super().__init__(
            parent,
            width=size,
            height=size,
            bg=bg,
            highlightthickness=0,
            cursor="hand2",
            **kwargs,
        )
        self._enabled = enabled
        self._command = command
        self._color   = color
        self._size    = size
        self._pad     = pad
        self._cache: dict[str, ImageTk.PhotoImage] = {}

        self._draw()
        self.bind("<Button-1>", self._on_click)

    # ── Rendering ─────────────────────────────────────────────────────────────

    def _render_circle(self, filled: bool) -> ImageTk.PhotoImage:
        """Return a cached PhotoImage for the given state."""
        key = "filled" if filled else "outline"
        if key in self._cache:
            return self._cache[key]

        # Draw at 4× for smooth anti-aliasing, then downscale with LANCZOS
        scale    = 4
        big_size = self._size * scale
        big_pad  = self._pad  * scale

        img  = Image.new("RGBA", (big_size, big_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        bbox = [big_pad, big_pad, big_size - big_pad, big_size - big_pad]

        if filled:
            draw.ellipse(bbox, fill=self._color)
        else:
            draw.ellipse(bbox, outline=self._color, width=2 * scale)

        small = img.resize((self._size, self._size), Image.Resampling.LANCZOS)
        tk_img = ImageTk.PhotoImage(small)
        self._cache[key] = tk_img
        return tk_img

    def _draw(self):
        self.delete("all")
        center = self._size // 2
        image  = self._render_circle(self._enabled)
        self.create_image(center, center, image=image, anchor="center")

    # ── Interaction ───────────────────────────────────────────────────────────

    def _on_click(self, _event=None):
        self._enabled = not self._enabled
        self._draw()
        if self._command:
            self._command(self._enabled)

    # ── Public API ────────────────────────────────────────────────────────────

    def get(self) -> bool:
        """Return the current toggle state."""
        return self._enabled

    def set(self, value: bool) -> None:
        """Set state programmatically without firing the command callback."""
        self._enabled = bool(value)
        self._draw()

    def set_color(self, color: str) -> None:
        """Change the accent colour at runtime and redraw (clears image cache)."""
        self._color = color
        self._cache.clear()
        self._draw()
