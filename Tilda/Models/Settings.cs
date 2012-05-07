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

namespace Tilda.Models {
    static class Settings {

        // You may modify the following either during execution or here
        public static String outputPath = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + Path.DirectorySeparatorChar + "temp";
        public static String outputMediaPath = "assets";
        public static String outputMediaFullPath = outputPath + Path.DirectorySeparatorChar + outputMediaPath;
        private static Random rand = new Random(Int32.MaxValue);
        
        //you shouldn't modify anything below

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
            } catch(Exception e) {
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
            //we need some formula here. right now we support 16:9 and 4:3
            //TODO: find appropriate formula if possible
            /*float ratio = (float)PresentationWidth() / (float)PresentationHeight();
            if(ratio - 1.33333 < .001)
                return 2.0f;
            else if(ratio - 1.77777 < .001)
                return 2.0f;
            else
                return 1.6f;*/
            //for some reason 2.0 seems to be just working now regardless of resolution
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
