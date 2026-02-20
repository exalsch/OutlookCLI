using OutlookCLI.Models;
using Xunit;

namespace OutlookCLI.Tests;

public class CommandResultTests
{
    [Fact]
    public void Ok_ShouldCreateSuccessResult()
    {
        var result = CommandResult<string>.Ok("test command", "test data");

        Assert.True(result.Success);
        Assert.Equal("test command", result.Command);
        Assert.Equal("test data", result.Data);
        Assert.Null(result.Error);
    }

    [Fact]
    public void Fail_ShouldCreateErrorResult()
    {
        var result = CommandResult<string>.Fail("test command", "ERROR_CODE", "Error message");

        Assert.False(result.Success);
        Assert.Equal("test command", result.Command);
        Assert.Null(result.Data);
        Assert.NotNull(result.Error);
        Assert.Equal("ERROR_CODE", result.Error.Code);
        Assert.Equal("Error message", result.Error.Message);
    }

    [Fact]
    public void Ok_WithMetadata_ShouldIncludeMetadata()
    {
        var metadata = new ResultMetadata { Count = 5 };
        var result = CommandResult<string>.Ok("test", "data", metadata);

        Assert.NotNull(result.Metadata);
        Assert.Equal(5, result.Metadata.Count);
    }
}
