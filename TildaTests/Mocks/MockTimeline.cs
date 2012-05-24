using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks
{
    class MockTimeline : TimeLine
    {
        Sequence mainSequence = new MockSequence();
        public Application Application
        {
            get { return this.Application; }
        }

        public Sequences InteractiveSequences
        {
            get { throw new NotImplementedException(); }
        }

        public Sequence MainSequence
        {
            get { return this.mainSequence; }
        }

        public dynamic Parent
        {
            get { throw new NotImplementedException(); }
        }
    }
}
