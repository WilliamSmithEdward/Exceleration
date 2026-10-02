namespace Exceleration.Tests;

/// <summary>
/// How the constructor opens the file, and the copy option.
/// </summary>
public class FileAccessTests
{
    [Fact]
    public void Reads_a_workbook_another_process_holds_open_for_writing()
    {
        using var folder = new TempFolder();
        var path = Xlsx.Prices(folder.File("prices.xlsx"));

        // How Excel holds a workbook it has open: read and write access, others may read.
        using var held = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

        Assert.Equal("Apple", new Workbook(path)["Sheet1"]["A2"].Value);
    }

    [Fact]
    public void Reads_a_workbook_another_reader_holds_open()
    {
        using var folder = new TempFolder();
        var path = Xlsx.Prices(folder.File("prices.xlsx"));

        using var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

        Assert.Equal("Apple", new Workbook(path)["Sheet1"]["A2"].Value);
    }

    [Fact]
    public void Several_threads_read_the_same_workbook_at_once()
    {
        using var folder = new TempFolder();
        var path = Xlsx.Prices(folder.File("prices.xlsx"));

        var values = Enumerable.Range(0, 16).AsParallel().WithDegreeOfParallelism(8)
            .Select(_ => new Workbook(path)["Sheet1"]["A3"].Value)
            .ToList();

        Assert.All(values, v => Assert.Equal("Pear", v));
    }

    // The copy option writes into the application's folder, the test's output folder here,
    // so each test uses a file name of its own and deletes what it made.
    [Fact]
    public void The_copy_option_reads_a_workbook_that_is_already_in_the_folder_it_copies_to()
    {
        var path = Path.Combine(AppContext.BaseDirectory, $"in-place-{Guid.NewGuid():N}.xlsx");
        try
        {
            Xlsx.Prices(path);

            var wb = new Workbook(path, copyFileToExeDirectoryBeforeRead: true);

            Assert.Equal(path, wb.FilePath);
            Assert.Equal("Apple", wb["Sheet1"]["A2"].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_copy_option_copies_into_the_application_folder_and_reads_the_copy()
    {
        using var folder = new TempFolder();
        var name = $"copied-{Guid.NewGuid():N}.xlsx";
        var source = Xlsx.Prices(folder.File(name));
        var copy = Path.Combine(AppContext.BaseDirectory, name);
        try
        {
            var wb = new Workbook(source, copyFileToExeDirectoryBeforeRead: true);

            Assert.Equal(copy, wb.FilePath);
            Assert.True(File.Exists(copy));
            Assert.Equal(File.ReadAllBytes(source), File.ReadAllBytes(copy));
            Assert.Equal("Apple", wb["Sheet1"]["A2"].Value);
        }
        finally
        {
            File.Delete(copy);
        }
    }
}
