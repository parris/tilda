using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Tilda.Models;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.IO.Packaging;
using Ionic.Zip;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace Tilda.Models {
    static class Settings {

        // You may modify the following either during execution or here
        public static String outputPath = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + Path.DirectorySeparatorChar + "temp";
        public static String outputMediaPath = "assets";
        public static String outputMediaFullPath = outputPath + Path.DirectorySeparatorChar + outputMediaPath;
        private static Random rand = new Random(Int32.MaxValue);
        
        //you shouldn't modify anything below
        //for numbered lists
        public static String[] AToZ = Enumerable.Range((int)'A', 26).Select(value => ((char)value).ToString()).ToArray();
        public static String[] aToz = Enumerable.Range((int)'a', 26).Select(value => ((char)value).ToString()).ToArray();
        //Rather than computing, We'll just list, it is unlikely values will be longer than this
        public static String[] romanNumeralsLC = new String[26]{"i","ii","iii","iv","v","vi","vii","viii",
                "ix","x","xi","xii","xiii","xiv","xv","xvi","xvii","xviii","xix",
                "xx","xxi","xxii","xxiii","xxiv","xxv","xxvi"};
        public static String[] romanNumeralsUC = Array.ConvertAll<string, string>(Settings.romanNumeralsLC, delegate(string s) { return s.ToUpper(); });

        /**
         * @returns PowerPoint.Presention the current presentation object
         */
        public static PowerPoint.Presentation ActivePresentation() {
            try {
                return Globals.ThisAddIn.Application.ActiveWindow.Presentation;
            } catch(Exception e) {
                return new MockPresentation(); //debug mode/no preso/good luck kids
            }
        }

        /**
         * @returns PowerPoint.Slide the current slide object
         */
        public static PowerPoint.Slide ActiveSlide() {
            try {
                return Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            } catch {
                return new MockSlide();
            }
        }

        /**
         * @returns int Width of current slide
         */
        public static int PresentationWidth() {
            return (int)ActivePresentation().PageSetup.SlideWidth * 2;
        }

        /**
         * @returns int Height of current slide
         */
        public static int PresentationHeight() {
            return (int)ActivePresentation().PageSetup.SlideHeight * 2;
        }

        public static float Scaler() {
            return 2.0f; 
        }

        public static String NextRandomValue() {
            return rand.Next().ToString();
        }

        public static String PresoSettingsToJS() {
            return "var settings = {'width':"+PresentationWidth()+",'height':"+PresentationHeight()+"};";
        }
    }
}
