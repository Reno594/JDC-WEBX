using System;
using QuestPDF.Drawing;
using QuestPDF.Infrastructure;

namespace QuestPDF.Elements
{
    internal sealed class Extend : ContainerElement
    {
        public bool ExtendVertical { get; set; }
        public bool ExtendHorizontal { get; set; }
        
        internal override SpacePlan Measure(Size availableSpace)
        {
            var childSize = base.Measure(availableSpace);

            if (childSize.Type is SpacePlanType.Empty or SpacePlanType.Wrap)
                return childSize;
            
            var targetSize = GetTargetSize(availableSpace, childSize);
            
            if (childSize.Type == SpacePlanType.PartialRender)
                return SpacePlan.PartialRender(targetSize);
            
            if (childSize.Type == SpacePlanType.FullRender)
                return SpacePlan.FullRender(targetSize);
            
            throw new NotSupportedException();
        }

        private Size GetTargetSize(Size availableSpace, Size childSize)
        {
            return new Size(
                ExtendHorizontal ? availableSpace.Width : childSize.Width, 
                ExtendVertical ? availableSpace.Height : childSize.Height);
        }
        
        internal override string? GetCompanionHint()
        {
            return (ExtendVertical, ExtendHorizontal) switch
            {
                (true, true) => "Both axes",
                (true, false) => "Vertical axis",
                (false, true) => "Horizontal axis",
                (false, false) => null
            };
        }
    }
}