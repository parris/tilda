using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Tilda.Models;
using System.Windows.Forms;

namespace Tilda {
    public partial class TildaRibbon {
        private void TildaRibbon_Load(object sender, RibbonUIEventArgs e) {

        }

        private void exportTildaSlide_Click(object sender, RibbonControlEventArgs e) {
            PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            //MediaElementUpload uploader = new MediaElementUpload("Narration");
            //uploader.exportFile();
            TildaSlide slide = new TildaSlide(currentSlide);
            MessageBox.Show(slide.exportSlide());
        }
    }
}
