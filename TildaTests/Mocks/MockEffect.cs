using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks {
    class MockEffect : Effect {

        private MockTiming timing;

        public MockEffect() {
            timing = new MockTiming();
        }

        public Application Application {
            get { throw new NotImplementedException(); }
        }

        public AnimationBehaviors Behaviors {
            get { throw new NotImplementedException(); }
        }

        public void Delete() {
            throw new NotImplementedException();
        }

        public string DisplayName {
            get { throw new NotImplementedException(); }
        }

        public EffectInformation EffectInformation {
            get { throw new NotImplementedException(); }
        }

        public EffectParameters EffectParameters {
            get { throw new NotImplementedException(); }
        }

        public MsoAnimEffect EffectType {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoTriState Exit {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public int Index {
            get { throw new NotImplementedException(); }
        }

        public void MoveAfter(Effect Effect) {
            throw new NotImplementedException();
        }

        public void MoveBefore(Effect Effect) {
            throw new NotImplementedException();
        }

        public void MoveTo(int toPos) {
            throw new NotImplementedException();
        }

        public int Paragraph {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public Shape Shape {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public int TextRangeLength {
            get { throw new NotImplementedException(); }
        }

        public int TextRangeStart {
            get { throw new NotImplementedException(); }
        }

        public Timing Timing {
            get { return this.timing; }
        }
    }
}
