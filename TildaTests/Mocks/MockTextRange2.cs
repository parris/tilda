using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TildaTests.Mocks {
    class MockTextRange2 : TextRange2{
        //private MockTextRange2 paragraphs = new MockTextRange2();
        private List<TextRange2> pgs = new List<TextRange2>();
        private String content = "";
        private ParagraphFormat2 pgformat = new MockParagraphFormat2();

        public MockTextRange2(){
        }

        public MockTextRange2(String content){
            this.content = content;
        }

        public void AddPeriods() {
            throw new NotImplementedException();
        }

        public dynamic Application {
            get { throw new NotImplementedException(); }
        }

        public float BoundHeight {
            get { throw new NotImplementedException(); }
        }

        public float BoundLeft {
            get { throw new NotImplementedException(); }
        }

        public float BoundTop {
            get { throw new NotImplementedException(); }
        }

        public float BoundWidth {
            get { throw new NotImplementedException(); }
        }

        public void ChangeCase(MsoTextChangeCase Type) {
            throw new NotImplementedException();
        }

        public void Copy() {
            throw new NotImplementedException();
        }

        public int Count {
            get { throw new NotImplementedException(); }
        }

        public int Creator {
            get { throw new NotImplementedException(); }
        }

        public void Cut() {
            throw new NotImplementedException();
        }

        public void Delete() {
            throw new NotImplementedException();
        }

        public TextRange2 Find(string FindWhat, int After = 0, MsoTriState MatchCase = MsoTriState.msoFalse, MsoTriState WholeWords = MsoTriState.msoFalse) {
            throw new NotImplementedException();
        }

        public Font2 Font {
            get { throw new NotImplementedException(); }
        }

        public System.Collections.IEnumerator GetEnumerator() {
            return this.pgs.GetEnumerator();
        }

        public TextRange2 InsertAfter(string NewText = "") {
            throw new NotImplementedException();
        }

        public TextRange2 InsertBefore(string NewText = "") {
            throw new NotImplementedException();
        }

        public TextRange2 InsertSymbol(string FontName, int CharNumber, MsoTriState Unicode = MsoTriState.msoFalse) {
            throw new NotImplementedException();
        }

        public TextRange2 Item(object Index) {
            throw new NotImplementedException();
        }

        public MsoLanguageID LanguageID {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public int Length {
            get { throw new NotImplementedException(); }
        }

        public void LtrRun() {
            throw new NotImplementedException();
        }

        public ParagraphFormat2 ParagraphFormat {
            get { return this.pgformat; }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public TextRange2 Paste() {
            throw new NotImplementedException();
        }

        public TextRange2 PasteSpecial(MsoClipboardFormat Format) {
            throw new NotImplementedException();
        }

        public void RemovePeriods() {
            throw new NotImplementedException();
        }

        public TextRange2 Replace(string FindWhat, string ReplaceWhat, int After = 0, MsoTriState MatchCase = MsoTriState.msoFalse, MsoTriState WholeWords = MsoTriState.msoFalse) {
            throw new NotImplementedException();
        }

        public void RotatedBounds(out float X1, out float Y1, out float X2, out float Y2, out float X3, out float Y3, out float x4, out float y4) {
            throw new NotImplementedException();
        }

        public void RtlRun() {
            throw new NotImplementedException();
        }

        public void Select() {
            throw new NotImplementedException();
        }

        public int Start {
            get { throw new NotImplementedException(); }
        }

        public string Text {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public TextRange2 TrimText() {
            throw new NotImplementedException();
        }

        public TextRange2 get_Characters(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }

        public TextRange2 get_Lines(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }

        public TextRange2 get_MathZones(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }

        public TextRange2 get_Paragraphs(int Start = -1, int Length = -1) {
            return new MockTextRange2();
        }

        public TextRange2 get_Runs(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }

        public TextRange2 get_Sentences(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }

        public TextRange2 get_Words(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }

        public void set_Paragraphs(List<MockTextRange2> pgraphs) {
            
        }
    }
}
