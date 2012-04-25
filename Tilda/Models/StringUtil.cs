using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace Tilda.Models {
    static class StringUtil {
        public static string Repeat(string input, int count) {
            StringBuilder builder = new StringBuilder(
                (input == null ? 0 : input.Length) * count);

            for (int i = 0; i < count; i++) builder.Append(input);

            return builder.ToString();
        }

        public static int CountStringOccurrences(string text, string pattern, int start = 0, int end = 0) {
            // Loop through all instances of the string 'text'.
            if (end == 0)
                end = text.Length;

            int count = 0;
            int i = start;

            while ((i = text.IndexOf(pattern, i)) != -1) {
                i += pattern.Length;
                count++;
            }
            return count;
        }
    }
}
