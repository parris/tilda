using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace Tilda.Models {
    class MockSlide : Slide {
        public Application Application {
            get { throw new NotImplementedException(); }
        }

        public void ApplyTemplate(string FileName) {
            throw new NotImplementedException();
        }

        public void ApplyTheme(string themeName) {
            throw new NotImplementedException();
        }

        public void ApplyThemeColorScheme(string themeColorSchemeName) {
            throw new NotImplementedException();
        }

        public ShapeRange Background {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoBackgroundStyleIndex BackgroundStyle {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public ColorScheme ColorScheme {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Comments Comments {
            get { throw new NotImplementedException(); }
        }

        public void Copy() {
            throw new NotImplementedException();
        }

        public CustomLayout CustomLayout {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public CustomerData CustomerData {
            get { throw new NotImplementedException(); }
        }

        public void Cut() {
            throw new NotImplementedException();
        }

        public void Delete() {
            throw new NotImplementedException();
        }

        public Design Design {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoTriState DisplayMasterShapes {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public SlideRange Duplicate() {
            throw new NotImplementedException();
        }

        public void Export(string FileName, string FilterName, int ScaleWidth = 0, int ScaleHeight = 0) {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.MsoTriState FollowMasterBackground {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoTriState HasNotesPage {
            get { throw new NotImplementedException(); }
        }

        public HeadersFooters HeadersFooters {
            get { throw new NotImplementedException(); }
        }

        public Hyperlinks Hyperlinks {
            get { throw new NotImplementedException(); }
        }

        public PpSlideLayout Layout {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Master Master {
            get { throw new NotImplementedException(); }
        }

        public void MoveTo(int toPos) {
            throw new NotImplementedException();
        }

        public void MoveToSectionStart(int toSection) {
            throw new NotImplementedException();
        }

        public string Name {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public SlideRange NotesPage {
            get { throw new NotImplementedException(); }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public int PrintSteps {
            get { throw new NotImplementedException(); }
        }

        public void PublishSlides(string SlideLibraryUrl, bool Overwrite = false, bool UseSlideOrder = false) {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Core.Scripts Scripts {
            get { throw new NotImplementedException(); }
        }

        public int SectionNumber {
            get { throw new NotImplementedException(); }
        }

        public void Select() {
            throw new NotImplementedException();
        }

        public Shapes Shapes {
            get { throw new NotImplementedException(); }
        }

        public int SlideID {
            get { throw new NotImplementedException(); }
        }

        public int SlideIndex {
            get { throw new NotImplementedException(); }
        }

        public int SlideNumber {
            get { throw new NotImplementedException(); }
        }

        public SlideShowTransition SlideShowTransition {
            get { throw new NotImplementedException(); }
        }

        public Tags Tags {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.ThemeColorScheme ThemeColorScheme {
            get { throw new NotImplementedException(); }
        }

        public TimeLine TimeLine {
            get { throw new NotImplementedException(); }
        }

        public int sectionIndex {
            get { throw new NotImplementedException(); }
        }
    }
}
