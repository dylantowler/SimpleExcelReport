namespace ExcelReport.Tests
{
    using NUnit.Framework;

    namespace ExcelReport.Tests
    {
        [TestFixture]
        public class Extensions
        {
            [Test]
            public void TestContiguous()
            {
                Assert.True(new[] { 1, 2, 3 }.Contiguous());

                Assert.False(new[] { 1, 2, 4 }.Contiguous());

                Assert.True(new[] { 0, 1, 2 }.Contiguous());

                Assert.False(new[] { 0, 2, 3 }.Contiguous());

                Assert.True(new[] { 3, 1, 2 }.Contiguous());

                Assert.True(new[] { -1, -2, -3 }.Contiguous());

                Assert.False(new[] { -1, -2, -4 }.Contiguous());

                Assert.True(new[] { 0, -1, -2 }.Contiguous());

                Assert.False(new[] { 0, -2, -3 }.Contiguous());

                Assert.True(new[] { -3, -1, -2 }.Contiguous());

                Assert.True(new[] { -3, -2, -1, 0, 1, 2, 3 }.Contiguous());

                Assert.False(new[] { -3, -2, -1, 0, 1, 2, 4 }.Contiguous());

                Assert.False(new[] { -4, -2, -1, 0, 1, 2, 3 }.Contiguous());
            }
        }
    }
}
