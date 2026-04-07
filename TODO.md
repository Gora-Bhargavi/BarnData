# BarnData Build Fix for .NET 10 - TODO

## Status: In Progress

### Step 1: [PENDING] Create TODO.md (current)
### Step 2: [COMPLETED ✓] Fix IsTagDuplicateAsync signature mismatch
   - Update AnimalService.cs / IAnimalService.cs to add 4-param overload with DateTime? killDate
   - Or fix AnimalController CheckTag call

### Step 3: [IN PROGRESS] Fix Razor ToString() errors in 7 views
   - Animal/Detail.cshtml [DONE]
   - Animal/Edit.cshtml, Index.cshtml
   - Report/Tally.cshtml, TallyPrint.cshtml, TallyToday.cshtml, VendorAnimals.cshtml, VendorAnimalsPrint.cshtml
   - Animal/Detail.cshtml, Edit.cshtml, Index.cshtml
   - Report/Tally.cshtml, TallyPrint.cshtml, TallyToday.cshtml, VendorAnimals.cshtml, VendorAnimalsPrint.cshtml
   - Replace .ToString(format) → interpolated $"{expr:format}"

### Step 4: [PENDING] Fix ImportController ?? operator errors (lines 72,92)
### Step 5: [PENDING] Fix nullable warnings (ReportController, views)
### Step 6: [PENDING] Update csproj TFMs to net10.0 (optional)
### Step 7: [PENDING] dotnet restore &amp;&amp; dotnet build -- verify clean build
### Step 8: [PENDING] dotnet run -- test app
### Step 9: [COMPLETED] attempt_completion

**Next:** Step 3
