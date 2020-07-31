using System;
using Xunit;

namespace DTE_ComparaArco
{
    public class TestComparaArco
    {

        //TODO: Agregar Test Unitarios
        [Fact]
        public void TestCuadrado()
        {
            //Arrange
            DTE_Compara xxx = new DTE_Compara();
            //Act
            int v = DTE_Compara.cuadrado(2);
            //Assert
            Assert.Equal(4, v);
        }
    }
}
