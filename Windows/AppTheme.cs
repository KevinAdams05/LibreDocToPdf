using System.Drawing;
using System.Windows.Forms;

namespace LibreDocToPdf
{
    public class AppTheme
    {
        public Color FormBack { get; init; }
        public Color FormFore { get; init; }
        public Color ControlBack { get; init; }
        public Color ControlFore { get; init; }
        public Color MenuBack { get; init; }
        public Color MenuFore { get; init; }
        public Color ButtonBack { get; init; }
        public Color ButtonFore { get; init; }
        public Color ButtonBorder { get; init; }
        public Color AccentBack { get; init; }
        public Color AccentFore { get; init; }

        public static AppTheme Light { get; } = new()
        {
            FormBack = Color.FromArgb(249, 249, 249),
            FormFore = Color.FromArgb(30, 30, 30),
            ControlBack = Color.White,
            ControlFore = Color.FromArgb(30, 30, 30),
            MenuBack = Color.FromArgb(243, 243, 243),
            MenuFore = Color.FromArgb(30, 30, 30),
            ButtonBack = Color.FromArgb(251, 251, 251),
            ButtonFore = Color.FromArgb(30, 30, 30),
            ButtonBorder = Color.FromArgb(180, 180, 180),
            AccentBack = Color.FromArgb(0, 120, 212),
            AccentFore = Color.White,
        };

        public static AppTheme Dark { get; } = new()
        {
            FormBack = Color.FromArgb(32, 32, 32),
            FormFore = Color.FromArgb(230, 230, 230),
            ControlBack = Color.FromArgb(45, 45, 45),
            ControlFore = Color.FromArgb(230, 230, 230),
            MenuBack = Color.FromArgb(45, 45, 45),
            MenuFore = Color.FromArgb(230, 230, 230),
            ButtonBack = Color.FromArgb(55, 55, 55),
            ButtonFore = Color.FromArgb(230, 230, 230),
            ButtonBorder = Color.FromArgb(80, 80, 80),
            AccentBack = Color.FromArgb(0, 120, 212),
            AccentFore = Color.White,
        };

        public ToolStripRenderer GetMenuRenderer()
        {
            return new ToolStripProfessionalRenderer(new ThemedColorTable(this));
        }

        private class ThemedColorTable : ProfessionalColorTable
        {
            private readonly AppTheme _theme;

            public ThemedColorTable(AppTheme theme) => _theme = theme;

            public override Color MenuItemSelected => Blend(_theme.MenuBack, Color.White, 0.15f);
            public override Color MenuItemBorder => _theme.ButtonBorder;
            public override Color MenuStripGradientBegin => _theme.MenuBack;
            public override Color MenuStripGradientEnd => _theme.MenuBack;
            public override Color ToolStripDropDownBackground => _theme.MenuBack;
            public override Color ImageMarginGradientBegin => _theme.MenuBack;
            public override Color ImageMarginGradientMiddle => _theme.MenuBack;
            public override Color ImageMarginGradientEnd => _theme.MenuBack;
            public override Color MenuItemPressedGradientBegin => Blend(_theme.MenuBack, Color.White, 0.1f);
            public override Color MenuItemPressedGradientEnd => Blend(_theme.MenuBack, Color.White, 0.1f);
            public override Color MenuItemSelectedGradientBegin => Blend(_theme.MenuBack, Color.White, 0.15f);
            public override Color MenuItemSelectedGradientEnd => Blend(_theme.MenuBack, Color.White, 0.15f);
            public override Color SeparatorDark => _theme.ButtonBorder;
            public override Color SeparatorLight => _theme.MenuBack;

            private static Color Blend(Color baseColor, Color blendColor, float amount)
            {
                int r = (int)(baseColor.R + (blendColor.R - baseColor.R) * amount);
                int g = (int)(baseColor.G + (blendColor.G - baseColor.G) * amount);
                int b = (int)(baseColor.B + (blendColor.B - baseColor.B) * amount);
                return Color.FromArgb(r, g, b);
            }
        }
    }
}
