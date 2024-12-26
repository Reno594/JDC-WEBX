﻿using NUnit.Framework;
using QuestPDF.Drawing;
using QuestPDF.Elements;
using QuestPDF.Infrastructure;
using QuestPDF.UnitTests.TestEngine;

namespace QuestPDF.UnitTests
{
    [TestFixture]
    public class PaddingTests
    {
        private Padding GetPadding(TestPlan plan)
        {
            return new Padding()
            {
                Top = 10,
                Right = 20,
                Bottom = 30,
                Left = 40,
                
                Child = plan.CreateChild()
            };
        }

        [Test]
        public void Measure_General_EnoughSpace()
        {
            TestPlan
                .For(GetPadding)
                .MeasureElement(new Size(400, 300))
                .ExpectChildMeasure(new Size(340, 260), SpacePlan.FullRender(140, 60))
                .CheckMeasureResult(SpacePlan.FullRender(200, 100));
        } 
        
        [Test]
        public void Measure_NotEnoughWidth()
        {
            TestPlan
                .For(GetPadding)
                .MeasureElement(new Size(50, 300))
                .ExpectChildMeasure(Size.Zero, SpacePlan.PartialRender(Size.Zero))
                .CheckMeasureResult(SpacePlan.Wrap("The available space is negative."));
        }
        
        [Test]
        public void Measure_NotEnoughHeight()
        {
            TestPlan
                .For(GetPadding)
                .MeasureElement(new Size(20, 300))
                .ExpectChildMeasure(Size.Zero, SpacePlan.PartialRender(Size.Zero))
                .CheckMeasureResult(SpacePlan.Wrap("The available space is negative."));
        }
        
        [Test]
        public void Measure_AcceptsPartialRender()
        {
            TestPlan
                .For(GetPadding)
                .MeasureElement(new Size(400, 300))
                .ExpectChildMeasure(new Size(340, 260), SpacePlan.PartialRender(40, 160))
                .CheckMeasureResult(SpacePlan.PartialRender(100, 200));
        }
        
        [Test]
        public void Measure_AcceptsWrap()
        {
            TestPlan
                .For(GetPadding)
                .MeasureElement(new Size(400, 300))
                .ExpectChildMeasure(new Size(340, 260), SpacePlan.Wrap("Mock"))
                .CheckMeasureResult(SpacePlan.Wrap("Forwarded from child"));
        }
        
        [Test]
        public void Draw_General()
        {
            TestPlan
                .For(GetPadding)
                .DrawElement(new Size(400, 300))
                .ExpectCanvasTranslate(new Position(40, 10))
                .ExpectChildDraw(new Size(340, 260))
                .ExpectCanvasTranslate(new Position(-40, -10))
                .CheckDrawResult();
        }
    }
}