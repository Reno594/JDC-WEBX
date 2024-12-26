﻿using System.Linq;
using QuestPDF.Drawing;
using QuestPDF.Infrastructure;

namespace QuestPDF.Elements
{
    internal sealed class ScaleToFit : ContainerElement
    {
        internal override SpacePlan Measure(Size availableSpace)
        {
            var perfectScale = FindPerfectScale(availableSpace);

            if (perfectScale == null)
                return SpacePlan.Wrap("Cannot find the perfect scale to fit the child element in the available space.");

            var scaledSpace = ScaleSize(availableSpace, 1 / perfectScale.Value);
            var childSizeInScale = base.Measure(scaledSpace);
            var childSizeInOriginalScale = ScaleSize(childSizeInScale, perfectScale.Value);
            return SpacePlan.FullRender(childSizeInOriginalScale);
        }
        
        internal override void Draw(Size availableSpace)
        {
            var perfectScale = FindPerfectScale(availableSpace);
            
            if (!perfectScale.HasValue)
                return;

            var targetScale = perfectScale.Value;
            var targetSpace = ScaleSize(availableSpace, 1 / targetScale);
            
            Canvas.Scale(targetScale, targetScale);
            Child?.Draw(targetSpace);
            Canvas.Scale(1 / targetScale, 1 / targetScale);
        }

        private static Size ScaleSize(Size size, float factor)
        {
            return new Size(size.Width * factor, size.Height * factor);
        }
        
        private float? FindPerfectScale(Size availableSpace)
        {
            if (ChildFits(1))
                return 1;
            
            var maxScale = 1f;
            var minScale = Size.Epsilon;

            var lastWorkingScale = (float?)null;
            
            foreach (var _ in Enumerable.Range(0, 8))
            {
                var halfScale = (maxScale + minScale) / 2;

                if (ChildFits(halfScale))
                {
                    minScale = halfScale;
                    lastWorkingScale = halfScale;
                }
                else
                {
                    maxScale = halfScale;
                }
            }
            
            return lastWorkingScale;
            
            bool ChildFits(float scale)
            {
                var scaledSpace = ScaleSize(availableSpace, 1 / scale);
                return base.Measure(scaledSpace).Type is SpacePlanType.Empty or SpacePlanType.FullRender;
            }
        }
    }
}