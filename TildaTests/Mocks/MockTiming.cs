using Tilda.Models;
using TildaTests.Mocks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks {
    class MockTiming : Timing {

        private float duration = 0f;
        private float triggerDelayTime = 0f;

        public float Accelerate {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Application Application {
            get { throw new NotImplementedException(); }
        }

        public Microsoft.Office.Core.MsoTriState AutoReverse {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoTriState BounceEnd {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float BounceEndIntensity {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float Decelerate {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float Duration {
            get {
                return this.duration;
            }
            set {
                this.duration = value;
            }
        }

        public dynamic Parent {
            get { throw new NotImplementedException(); }
        }

        public int RepeatCount {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float RepeatDuration {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoAnimEffectRestart Restart {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoTriState RewindAtEnd {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoTriState SmoothEnd {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.MsoTriState SmoothStart {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float Speed {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public string TriggerBookmark {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public float TriggerDelayTime {
            get {
                return this.triggerDelayTime;
            }
            set {
                this.triggerDelayTime = value;
            }
        }

        public Shape TriggerShape {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }

        public MsoAnimTriggerType TriggerType {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }
    }
}
