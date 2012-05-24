using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks {

    [Serializable]
    class MockTextRange2 : TextRange2{
        private List<TextRange2> trList;
        private String text = "";
        private Font2 font = new MockFont();
        private MockTextFrame2 parent;
        private ParagraphFormat2 pgformat = new MockParagraphFormat2();

        public MockTextRange2(){
        }

        public MockTextRange2(String content, MockTextFrame2 parent = null){
            this.text = content;
            this.parent = parent;
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
            get { return this.font; }
        }

        public System.Collections.IEnumerator GetEnumerator() {
            return this.trList.GetEnumerator();
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
            get { return this.parent; }
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
                return this.text;
            }
            set {
                this.text = value;
            }
        }

        public TextRange2 TrimText() {
            throw new NotImplementedException();
        }

        public TextRange2 get_Characters(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }


        public TextRange2 get_MathZones(int Start = -1, int Length = -1) {
            throw new NotImplementedException();
        }

        public TextRange2 get_Lines(int Start = -1, int Length = -1) {
            return this.getLinesOrPgs('~');
        }

        public TextRange2 get_Paragraphs(int Start = -1, int Length = -1) {
            return this.getLinesOrPgs('\r');
        }

        private TextRange2 getLinesOrPgs(char splitChar) {
            List<TextRange2> result = new List<TextRange2>();
            this.trList = null;
            foreach(String t in this.text.Split(splitChar).ToList<String>()) {
                MockTextRange2 tr = MockHelper.DeepClone(this);
                tr.Text = t;
                this.modTextRange(tr);
                result.Add(tr);
            }

            this.trList = result;
            return MockHelper.DeepClone(this);
        }

        private void modTextRange(MockTextRange2 tr) {
            if(tr.Text.Length == 0)
                return;
            if(tr.Text[0]=='`') {
                // first level bullet
                tr.ParagraphFormat.Bullet.Type = MsoBulletType.msoBulletUnnumbered;
                tr.ParagraphFormat.Bullet.Character = 167; // square
                tr.ParagraphFormat.IndentLevel = 1;
                tr.Text = tr.Text.Substring(1);
            } else if(tr.Text[0] == '^') {
                // second level bullet
                tr.ParagraphFormat.Bullet.Type = MsoBulletType.msoBulletUnnumbered;
                tr.ParagraphFormat.Bullet.Character = 167; // square
                tr.ParagraphFormat.IndentLevel = 2;
                tr.ParagraphFormat.LeftIndent = tr.ParagraphFormat.LeftIndent * 2;
                tr.ParagraphFormat.FirstLineIndent = tr.ParagraphFormat.FirstLineIndent * 2;
                tr.Text = tr.Text.Substring(1);
            } else if(tr.Text[0] == '*') {
                // third level bullet
                tr.ParagraphFormat.Bullet.Type = MsoBulletType.msoBulletUnnumbered;
                tr.ParagraphFormat.Bullet.Character = 167; // square
                tr.ParagraphFormat.IndentLevel = 3;
                tr.ParagraphFormat.LeftIndent = tr.ParagraphFormat.LeftIndent * 2;
                tr.ParagraphFormat.FirstLineIndent = tr.ParagraphFormat.FirstLineIndent * 2;
                tr.Text = tr.Text.Substring(1);
            }
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
    }
}
