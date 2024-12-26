﻿using QuestPDF.Drawing;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace QuestPDF.Elements
{
    internal sealed class PageBreak : Element, IStateful
    {
        internal override SpacePlan Measure(Size availableSpace)
        {
            if (availableSpace.IsNegative())
                return SpacePlan.Wrap("The available space is negative.");

            if (IsRendered)
                return SpacePlan.Empty();

            return SpacePlan.PartialRender(Size.Zero);
        }

        internal override void Draw(Size availableSpace)
        {
            IsRendered = true;
        }
        
        #region IStateful
        
        private bool IsRendered { get; set; }
    
        public void ResetState(bool hardReset = false) => IsRendered = false;
        public object GetState() => IsRendered;
        public void SetState(object state) => IsRendered = (bool) state;
    
        #endregion
    }
}