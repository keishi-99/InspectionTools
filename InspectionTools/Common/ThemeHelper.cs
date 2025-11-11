using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace InspectionTools.Common {
    public class ThemeHelper {
        /// <summary>
        /// BundledThemeの設定
        /// </summary>
        public static void SetBundledTheme(BundledTheme bundledTheme) {
            System.Windows.Application.Current.Resources.MergedDictionaries.Add(bundledTheme);
        }

        /// <summary>
        /// BundledThemeの設定
        /// </summary>
        public static void SetBundledTheme(BaseTheme baseTheme, PrimaryColor primaryColor, SecondaryColor secondaryColor) {
            SetBundledTheme(new BundledTheme {
                BaseTheme = baseTheme,
                PrimaryColor = primaryColor,
                SecondaryColor = secondaryColor
            });
        }

        /// <summary>
        /// BaseThemeの取得
        /// </summary>
        public static BaseTheme GetBaseTheme() {
            return GetBundledTheme()?.BaseTheme ?? BaseTheme.Light;
        }

        /// <summary>
        /// PrimaryColorの取得
        /// </summary>
        public static PrimaryColor GetPrimaryColor() {
            return GetBundledTheme()?.PrimaryColor ?? PrimaryColor.Blue;
        }

        /// <summary>
        /// SecondaryColorの取得
        /// </summary>
        public static SecondaryColor GetSecondaryColor() {
            return GetBundledTheme()?.SecondaryColor ?? SecondaryColor.Blue;
        }

        /// <summary>
        /// BundledThemeの取得
        /// </summary>
        public static BundledTheme GetBundledTheme() {
            var bundledTheme = System.Windows.Application.Current.Resources.MergedDictionaries
                                  .OfType<BundledTheme>()
                                  .FirstOrDefault();
            return bundledTheme ?? new BundledTheme() {
                BaseTheme = BaseTheme.Light,
                PrimaryColor = PrimaryColor.Green,
                SecondaryColor = SecondaryColor.Green
            };
        }
    }
}
