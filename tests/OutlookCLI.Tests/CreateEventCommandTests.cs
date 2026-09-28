using OutlookCLI.Commands.Calendar;
using Xunit;

namespace OutlookCLI.Tests;

public class CreateEventCommandTests
{
    [Fact]
    public void SplitAddresses_ShouldReturnEmpty_WhenNull()
    {
        Assert.Empty(CreateEventCommand.SplitAddresses(null));
    }

    [Fact]
    public void SplitAddresses_ShouldFlattenSpaceCommaAndSemicolonSeparated()
    {
        var result = CreateEventCommand.SplitAddresses(["a@x.com,b@x.com", "c@x.com; d@x.com", "e@x.com"]);

        Assert.Equal(new[] { "a@x.com", "b@x.com", "c@x.com", "d@x.com", "e@x.com" }, result);
    }

    [Fact]
    public void SplitAddresses_ShouldDropEmptyEntries()
    {
        var result = CreateEventCommand.SplitAddresses(["a@x.com,,", " ; "]);

        Assert.Equal(new[] { "a@x.com" }, result);
    }
}
