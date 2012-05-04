using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace TildaTests.Mocks
{
    class MockTimeline : TimeLine
    {

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
            get { throw new NotImplementedException(); }
        }

        public dynamic Parent
        {
            get { throw new NotImplementedException(); }
        }
    }
}
