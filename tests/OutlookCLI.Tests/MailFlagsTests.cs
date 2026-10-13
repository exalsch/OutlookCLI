using System.Text.Json;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;
using Xunit;

namespace OutlookCLI.Tests;

public class MailFlagsTests
{
    [Theory]
    [InlineData(0, MailFlagStatus.None)]
    [InlineData(1, MailFlagStatus.Complete)]
    [InlineData(2, MailFlagStatus.Flagged)]
    [InlineData(99, MailFlagStatus.None)]
    public void FromOutlook_ShouldMapOlFlagStatus(int raw, MailFlagStatus expected)
    {
        Assert.Equal(expected, MailFlags.FromOutlook(raw));
    }

    [Theory]
    [InlineData(MailFlagStatus.None, "none")]
    [InlineData(MailFlagStatus.Flagged, "flagged")]
    [InlineData(MailFlagStatus.Complete, "complete")]
    public void ToName_ShouldReturnApiName(MailFlagStatus status, string expected)
    {
        Assert.Equal(expected, MailFlags.ToName(status));
    }

    [Fact]
    public void NormalizeDate_ShouldMapOutlookNoneDateToNull()
    {
        Assert.Null(MailFlags.NormalizeDate(new DateTime(4501, 1, 1)));
        Assert.Null(MailFlags.NormalizeDate(DateTime.MinValue));
        Assert.Null(MailFlags.NormalizeDate(null));

        var real = new DateTime(2026, 10, 12);
        Assert.Equal(real, MailFlags.NormalizeDate(real));
    }

    [Theory]
    [InlineData(null, null)]
    [InlineData("", null)]
    [InlineData("   ", null)]
    [InlineData("Follow up", "Follow up")]
    public void NormalizeRequest_ShouldMapEmptyToNull(string? input, string? expected)
    {
        Assert.Equal(expected, MailFlags.NormalizeRequest(input));
    }

    [Fact]
    public void TryResolveFilter_NoOptions_ShouldReturnNoFilter()
    {
        Assert.True(MailFlags.TryResolveFilter(false, null, out var filter, out var error));
        Assert.Null(filter);
        Assert.Null(error);
    }

    [Fact]
    public void TryResolveFilter_Flagged_ShouldReturnFlagged()
    {
        Assert.True(MailFlags.TryResolveFilter(true, null, out var filter, out _));
        Assert.Equal(MailFlagStatus.Flagged, filter);
    }

    [Theory]
    [InlineData("none", MailFlagStatus.None)]
    [InlineData("complete", MailFlagStatus.Complete)]
    [InlineData("FLAGGED", MailFlagStatus.Flagged)]
    public void TryResolveFilter_FlagStatus_ShouldParse(string value, MailFlagStatus expected)
    {
        Assert.True(MailFlags.TryResolveFilter(false, value, out var filter, out _));
        Assert.Equal(expected, filter);
    }

    [Fact]
    public void TryResolveFilter_ConflictingOptions_ShouldFail()
    {
        Assert.False(MailFlags.TryResolveFilter(true, "complete", out _, out var error));
        Assert.NotNull(error);
    }

    [Fact]
    public void TryResolveFilter_UnknownStatus_ShouldFail()
    {
        Assert.False(MailFlags.TryResolveFilter(false, "marked", out _, out var error));
        Assert.NotNull(error);
    }

    [Fact]
    public void GetRestrictFilter_Flagged_ShouldAvoidEqualityComparison()
    {
        // "[FlagStatus] = 2" returns no items on Exchange cached-mode stores
        Assert.Equal("[FlagStatus] >= 2", MailFlags.GetRestrictFilter(MailFlagStatus.Flagged));
        Assert.Equal("[FlagStatus] = 1", MailFlags.GetRestrictFilter(MailFlagStatus.Complete));
        Assert.Equal("[FlagStatus] = 0", MailFlags.GetRestrictFilter(MailFlagStatus.None));
    }

    [Fact]
    public void JsonOutput_ShouldIncludeFlagFields()
    {
        var summary = new MailMessageSummary(
            "id", "Subject", "Sender", "sender@example.com", new DateTime(2026, 10, 9, 15, 47, 0),
            false, "Inbox", false, 0, new List<string>(), null,
            IsFlagged: true,
            FlagStatus: "flagged",
            FlagRequest: "Follow up",
            FlagDueDate: new DateTime(2026, 10, 12));

        var result = CommandResult<IEnumerable<MailMessageSummary>>.Ok("mail list", new[] { summary });
        var json = new JsonOutputFormatter().FormatMailList(result);

        using var doc = JsonDocument.Parse(json);
        var item = doc.RootElement.GetProperty("data")[0];
        Assert.True(item.GetProperty("isFlagged").GetBoolean());
        Assert.Equal("flagged", item.GetProperty("flagStatus").GetString());
        Assert.Equal("Follow up", item.GetProperty("flagRequest").GetString());
        Assert.Equal(new DateTime(2026, 10, 12), item.GetProperty("flagDueDate").GetDateTime());
        Assert.False(item.TryGetProperty("flagCompletedDate", out _)); // nulls are omitted
    }

    [Fact]
    public void JsonOutput_UnflaggedDefaults_ShouldReportNone()
    {
        var summary = new MailMessageSummary(
            "id", "Subject", "Sender", "sender@example.com", DateTime.Now,
            true, "Inbox", false, 0, new List<string>());

        var result = CommandResult<IEnumerable<MailMessageSummary>>.Ok("mail list", new[] { summary });
        var json = new JsonOutputFormatter().FormatMailList(result);

        using var doc = JsonDocument.Parse(json);
        var item = doc.RootElement.GetProperty("data")[0];
        Assert.False(item.GetProperty("isFlagged").GetBoolean());
        Assert.Equal("none", item.GetProperty("flagStatus").GetString());
    }
}
