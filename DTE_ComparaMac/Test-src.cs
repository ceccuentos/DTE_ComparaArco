using System;
using Xunit;

namespace DTE_ComparaArco
{
    public class TestComparaArco
    {

        [Fact]
        public void TestCuadrado()
        {
            //Arrange
            DTE_Compara xxx = new DTE_Compara();
            //Act
            //xxx.cuadrad
            int v = 0; //xxx.cuadrado(2);
            int zz_ = v;
            //Assert
            Assert.Equal(4, zz_);
        }
    }
}