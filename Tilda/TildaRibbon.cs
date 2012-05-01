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

namespace Tilda {
    public partial class TildaRibbon {
        private void TildaRibbon_Load(object sender, RibbonUIEventArgs e) {
        }

        private void exportTildaSlide_Click(object sender, RibbonControlEventArgs e) {
            setUpFolders();
            PowerPoint.Slide currentSlide = Settings.ActiveSlide();
            TildaSlide slide = new TildaSlide(currentSlide);
            export("preso.slides.push(" + slide.exportSlide() + ");");
            cleanUpFolders();
            MessageBox.Show("Saved :)");
        }

        private void exportTildaShape_Click(object sender, RibbonControlEventArgs e) {
            setUpFolders();
            String js = "";
            PowerPoint.PpSelectionType type = Globals.ThisAddIn.Application.ActiveWindow.Selection.Type;
            if (type == PowerPoint.PpSelectionType.ppSelectionNone ||
                type == PowerPoint.PpSelectionType.ppSelectionText)
                MessageBox.Show("You can only export slides right now via selection");
            else if(type == PowerPoint.PpSelectionType.ppSelectionSlides) {
                foreach(PowerPoint.Slide currentSlide in Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange) {
                    TildaSlide slide = new TildaSlide(currentSlide);
                    js += "preso.slides.push(" + slide.exportSlide() + ");";
                }
                export(js);
                MessageBox.Show("Saved :)");
            } else if(type == PowerPoint.PpSelectionType.ppSelectionShapes) {
                /*PowerPoint.Slide currentSlide = Settings.ActiveSlide();
                TildaSlide slide = new TildaSlide(currentSlide);
                slide.shapesRange = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                js += "preso.slides.push(" + slide.exportSlide() + ");";*/

                MessageBox.Show("You can only export slides right now via selection");
            }
            cleanUpFolders();
        }

        private void exportTildaPresentation_Click(object sender, RibbonControlEventArgs e) {
            PowerPoint.Presentation currentPreso = Settings.ActivePresentation();

            setUpFolders();
            String js = "";
            foreach(PowerPoint.Slide currentSlide in currentPreso.Slides) {
                TildaSlide slide = new TildaSlide(currentSlide);
                js += "preso.slides.push(" + slide.exportSlide() + ");";
            }
            export(js);
            cleanUpFolders();
            MessageBox.Show("Saved :)");
        }

        private void setUpFolders() {
            Directory.CreateDirectory(Settings.outputPath);
            Directory.CreateDirectory(Settings.outputMediaFullPath);
        }

        private void cleanUpFolders() {
            Directory.Delete(Settings.outputPath, true);
        }

        private void export(String js) {
            String path = Settings.outputPath;

            System.IO.File.WriteAllText(path + Path.DirectorySeparatorChar + "content.js", js);
            //zip it all up
            using (ZipFile zip = new ZipFile()) {
                //make zip file
                //add content
                zip.AddFile(path + Path.DirectorySeparatorChar + "content.js", "");
                zip.AddDirectory(path + Path.DirectorySeparatorChar + "assets", "assets");

                //add libs
                zip.AddDirectory(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + Path.DirectorySeparatorChar + "Web" + Path.DirectorySeparatorChar + "js", "js");
                zip.AddFile(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + Path.DirectorySeparatorChar + "Web" + Path.DirectorySeparatorChar + "index.html", "");

                zip.Save(path + Path.DirectorySeparatorChar + "slide.zip");
                //create dialog box
                SaveFileDialog dia = new SaveFileDialog();
                dia.Filter = "Zip File (*.zip)|*.zip|All files (*.*)|*.*";
                dia.FilterIndex = 2;
                dia.RestoreDirectory = true;

                //open dialog move file
                if (dia.ShowDialog() == DialogResult.OK)
                    File.Copy(path + Path.DirectorySeparatorChar + "slide.zip", dia.FileName, true);

                //remove zip file
                File.Delete(path + Path.DirectorySeparatorChar + "slide.zip");
            }
        }
    }
}
