using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks {
    class MockSequence : Sequence {

        private List<MockSequence> sequences;

        public MockSequence() {
            this.sequences = new List<MockSequence>();
        }

        public Effect AddEffect(Shape Shape, MsoAnimEffect effectId, MsoAnimateByLevel Level = MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType trigger = MsoAnimTriggerType.msoAnimTriggerOnPageClick, int Index = -1) {
            throw new NotImplementedException();
        }

        public Effect AddTriggerEffect(Shape pShape, MsoAnimEffect effectId, MsoAnimTriggerType trigger, Shape pTriggerShape, string bookmark = "", MsoAnimateByLevel Level = MsoAnimateByLevel.msoAnimateLevelNone) {
            throw new NotImplementedException();
        }

        public Application Application {
            get { throw new NotImplementedException(); }
        }

        public Effect Clone(Effect Effect, int Index = -1) {
            throw new NotImplementedException();
        }

        public Effect ConvertToAfterEffect(Effect Effect, MsoAnimAfterEffect After, int DimColor = 0, PpColorSchemeIndex DimSchemeColor = PpColorSchemeIndex.ppNotSchemeColor) {
            throw new NotImplementedException();
        }

        public Effect ConvertToAnimateBackground(Effect Effect, Microsoft.Office.Core.MsoTriState AnimateBackground) {
            throw new NotImplementedException();
        }

        public Effect ConvertToAnimateInReverse(Effect Effect, Microsoft.Office.Core.MsoTriState animateInReverse) {
            throw new NotImplementedException();
        }

        public Effect ConvertToBuildLevel(Effect Effect, MsoAnimateByLevel Level) {
            throw new NotImplementedException();
        }

        public Effect ConvertToTextUnitEffect(Effect Effect, MsoAnimTextUnitEffect unitEffect) {
            throw new NotImplementedException();
        }

        public int Count {
            get { throw new NotImplementedException(); }
        }

        public Effect FindFirstAnimationFor(Shape Shape) {
            throw new NotImplementedException();
        }

        public Effect FindFirstAnimationForClick(int click) {
            throw new NotImplementedException();
        }

        public System.Collections.IEnumerator GetEnumerator() {
            return sequences.GetEnumerator();
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public dynamic _Index(int Index) {
            throw new NotImplementedException();
        }

        public Effect this[int Index] {
            get { throw new NotImplementedException(); }
        }
    }
}
