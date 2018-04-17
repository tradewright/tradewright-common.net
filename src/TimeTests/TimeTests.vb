Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports TradeWright.Utilities.Time

<TestClass()> Public Class TimeTests

    <TestClass()> Public Class GetWeekNumberTests
        <TestMethod()> Public Sub TestGetWeekNumber10()
            Assert.AreEqual(15, GetWeekNumber(CDate("2018/4/12")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber20()
            Assert.AreEqual(1, GetWeekNumber(CDate("2018/1/1")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber30()
            Assert.AreEqual(1, GetWeekNumber(CDate("2018/1/7")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber40()
            Assert.AreEqual(2, GetWeekNumber(CDate("2018/1/8")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber50()
            Assert.AreEqual(53, GetWeekNumber(CDate("2018/12/31")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber60()
            Assert.AreEqual(52, GetWeekNumber(CDate("2018/12/30")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber70()
            Assert.AreEqual(0, GetWeekNumber(CDate("2017/1/1")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber80()
            Assert.AreEqual(1, GetWeekNumber(CDate("2017/1/2")))
        End Sub

        <TestMethod()> Public Sub TestGetWeekNumber90()
            Assert.AreEqual(52, GetWeekNumber(CDate("2017/12/31")))
        End Sub

    End Class

    <TestClass> Public Class GetWeekStartDateTests
        <TestMethod> Public Sub TestGetWeekStartDate10()
            Assert.AreEqual(CDate("2018/01/01"), GetWeekStartDate(CDate("2018/01/7")))
        End Sub

        <TestMethod> Public Sub TestGetWeekStartDate20()
            Assert.AreEqual(CDate("2018/04/09"), GetWeekStartDate(CDate("2018/04/14")))
        End Sub

    End Class

    <TestClass> Public Class GetWeekStartDateFromWeekNumberTests
        <TestMethod> Public Sub TestGetWeekStartDateFromWeekNumber10()
            Assert.AreEqual(CDate("2018/04/9"), GetWeekStartDateFromWeekNumber(15, CDate("2018/01/7")))
        End Sub

        <TestMethod> Public Sub TestGetWeekStartDateFromWeekNumber20()
            Assert.AreEqual(CDate("2018/12/31"), GetWeekStartDateFromWeekNumber(53, CDate("2018/01/7")))
        End Sub

        <TestMethod> Public Sub TestGetWeekStartDateFromWeekNumber30()
            Assert.AreEqual(CDate("2017/01/2"), GetWeekStartDateFromWeekNumber(1, CDate("2017/01/7")))
        End Sub

        <TestMethod> Public Sub TestGetWeekStartDateFromWeekNumber40()
            Assert.AreEqual(CDate("2017/12/25"), GetWeekStartDateFromWeekNumber(52, CDate("2017/01/7")))
        End Sub

        <TestMethod> Public Sub TestGetWeekStartDateFromWeekNumber50()
            Assert.AreEqual(CDate("2018/1/1"), GetWeekStartDateFromWeekNumber(53, CDate("2017/01/7")))
        End Sub

    End Class

    <TestClass> Public Class GetWorkingDayDateTests
        <TestMethod> Public Sub TestGetWorkingDayDate10()
            Assert.AreEqual(CDate("2018/01/1"), GetWorkingDayDate(1, CDate("2018/01/7")))
        End Sub

        <TestMethod> Public Sub TestGetWorkingDayDate20()
            Assert.AreEqual(CDate("2018/01/31"), GetWorkingDayDate(23, CDate("2018/01/7")))
        End Sub

        <TestMethod> Public Sub TestGetWorkingDayDate30()
            Assert.AreEqual(CDate("2018/05/2"), GetWorkingDayDate(88, CDate("2018/01/7")))
        End Sub

    End Class

    <TestClass> Public Class GetWorkingDayNumberTests
        <TestMethod> Public Sub TestGetWorkingDayNumber10()
            Assert.AreEqual(5, GetWorkingDayNumber(CDate("2018/01/7")))
        End Sub

    End Class

End Class