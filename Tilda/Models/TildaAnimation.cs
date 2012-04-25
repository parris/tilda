using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Tilda.Models
{
    class TildaAnimation
    {
        public PowerPoint.Effect effect;
        public TildaShape shape;
        public bool added = false;

        public TildaAnimation(PowerPoint.Effect effect, TildaShape shape)
        {
            this.effect = effect;
            this.shape = shape;
        }
    }
}
