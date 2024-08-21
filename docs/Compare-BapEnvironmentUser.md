﻿---
external help file: d365bap.tools-help.xml
Module Name: d365bap.tools
online version:
schema: 2.0.0
---

# Compare-BapEnvironmentUser

## SYNOPSIS
Compare the environment users

## SYNTAX

```
Compare-BapEnvironmentUser [-SourceEnvironmentId] <String> [-DestinationEnvironmentId] <String> [-ShowDiffOnly]
 [-IncludeAppIds] [-AsExcelOutput] [<CommonParameters>]
```

## DESCRIPTION
Enables the user to compare 2 x environments, with one as a source and the other as a destination

It will only look for users on the source, and use this as a baseline against the destination

## EXAMPLES

### EXAMPLE 1
```
Compare-BapEnvironmentD365App -SourceEnvironmentId eec2c11a-a4c7-4e1d-b8ed-f62acc9c74c6 -DestinationEnvironmentId 32c6b196-ef52-4c43-93cf-6ecba51e6aa1
```

This will get all system users from the Source Environment.
It will iterate over all of them, and validate against the Destination Environment.
It will exclude those with ApplicationId filled.

Sample output:
Email                          Name                           AppId                SourceId        DestinationId
-----                          ----                           -----                --------        -------------
aba@temp.com                   Austin Baker                                        f85bcd69-ef72-â€¦ 5aaac0ec-a91â€¦
ade@temp.com                   Alex Denver                                         39309a5c-7676-â€¦ 1d521227-43bâ€¦

### EXAMPLE 2
```
Compare-BapEnvironmentD365App -SourceEnvironmentId eec2c11a-a4c7-4e1d-b8ed-f62acc9c74c6 -DestinationEnvironmentId 32c6b196-ef52-4c43-93cf-6ecba51e6aa1 -IncludeAppIds
```

This will get all system users from the Source Environment.
It will iterate over all of them, and validate against the Destination Environment.
It will include those with ApplicationId filled.

Sample output:
Email                          Name                           AppId                SourceId        DestinationId
-----                          ----                           -----                --------        -------------
aba@temp.com                   Austin Baker                                        f85bcd69-ef72-â€¦ 5aaac0ec-a91â€¦
ade@temp.com                   Alex Denver                                         39309a5c-7676-â€¦ 1d521227-43bâ€¦
AIBuilder_StructuredML_Prod_Câ€¦ AIBuilder_StructuredML_Prod_Câ€¦ ff8a1ad8-a415-45c1-â€¦ 95dc9ca2-8185-â€¦ 328db0cc-14câ€¦
AIBuilderProd@onmicrosoft.com  AIBuilderProd, #               0a143f2d-2320-4141-â€¦ c96f82b8-320f-â€¦ 1831f4dc-4c5â€¦

### EXAMPLE 3
```
Compare-BapEnvironmentD365App -SourceEnvironmentId eec2c11a-a4c7-4e1d-b8ed-f62acc9c74c6 -DestinationEnvironmentId 32c6b196-ef52-4c43-93cf-6ecba51e6aa1 -IncludeAppIds -ShowDiffOnly
```

This will get all system users from the Source Environment.
It will iterate over all of them, and validate against the Destination Environment.
It will include those with ApplicationId filled.
It will only output the users that is missing in the destionation environment.

Sample output:
Email                          Name                           AppId                SourceId        DestinationId
-----                          ----                           -----                --------        -------------
d365-scm-operationdataserviceâ€¦ d365-scm-operationdataserviceâ€¦ 986556ed-a409-4339-â€¦ 5e077e6a-a0c9-â€¦ Missing
d365-scm-operationdataserviceâ€¦ d365-scm-operationdataserviceâ€¦ 14e80222-1878-455d-â€¦ 183ec023-9ccb-â€¦ Missing
def@temp.com                   Dustin Effect                                       01e37132-0a44-â€¦ Missing

### EXAMPLE 4
```
Compare-BapEnvironmentD365App -SourceEnvironmentId eec2c11a-a4c7-4e1d-b8ed-f62acc9c74c6 -DestinationEnvironmentId 32c6b196-ef52-4c43-93cf-6ecba51e6aa1 -AsExcelOutput
```

This will get all system users from the Source Environment.
It will iterate over all of them, and validate against the Destination Environment.
It will exclude those with ApplicationId filled.
Will output all details into an Excel file, that will auto open on your machine.

## PARAMETERS

### -SourceEnvironmentId
Environment Id of the source environment that you want to utilized as the baseline for the compare

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DestinationEnvironmentId
Environment Id of the destination environment that you want to validate against the baseline (source)

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowDiffOnly
Instruct the cmdlet to only output the differences that are not aligned between the source and destination

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeAppIds
Instruct the cmdlet to also include the users with the ApplicationId property filled

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsExcelOutput
Instruct the cmdlet to output all details directly to an Excel file

This makes it easier to deep dive into all the details returned from the API, and makes it possible for the user to persist the current state

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
Author: Mötz Jensen (@Splaxi)

## RELATED LINKS
