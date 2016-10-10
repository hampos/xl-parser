using FluentAssertions;
using Xunit;

namespace XlsxSaxExporter.Tests
{
    public class XlsxSheetDimensionsTests
    {
        [Fact]
        public void I_Can_Create_XlsxSheetDimensions_Without_Parameters_And_Dimensions_Will_Be_Zeroes()
        {
            var xlsxSheetDimensions = new XlsxSheetDimensions();

            xlsxSheetDimensions.Should().NotBeNull();
            xlsxSheetDimensions.MinRowNum.Should().Be(0);
            xlsxSheetDimensions.MaxRowNum.Should().Be(0);
            xlsxSheetDimensions.MinColNum.Should().Be(0);
            xlsxSheetDimensions.MaxColNum.Should().Be(0);
        }

        [Fact]
        public void Given_Row_And_Col_Nums_When_Creating_XlsxSheetDimensions_Then_XlsxSheetDimenions_Are_Created_With_Same_Nums()
        {
            var xlsxSheetDimensions = new XlsxSheetDimensions(1, 10, 1, 20);

            xlsxSheetDimensions.Should().NotBeNull();
            xlsxSheetDimensions.MinRowNum.Should().Be(1);
            xlsxSheetDimensions.MaxRowNum.Should().Be(10);
            xlsxSheetDimensions.MinColNum.Should().Be(1);
            xlsxSheetDimensions.MaxColNum.Should().Be(20);
        }
    }
}
