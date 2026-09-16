-- Compiled by the installed Word dictionary at runtime; no Word dependency at build time.
on run arguments
    if (count of arguments) is not 3 then error "Expected original, revised and output paths."
    set sourceFile to (POSIX file (item 1 of arguments)) as alias
    set revisedFile to (POSIX file (item 2 of arguments)) as alias
    set destinationFile to POSIX file (item 3 of arguments)
    set sourceDocument to missing value
    set comparisonDocument to missing value
    with timeout of 600 seconds
        tell application "Microsoft Word"
            activate
            try
                set sourceDocument to open file name sourceFile with read only and new
                set previousNames to name of every document
                compare sourceDocument path (revisedFile as text) author name "Redline"
                set resultCandidates to {}
                repeat with candidate in every document
                    if (name of candidate) is not in previousNames then
                        if (name of candidate) does not start with "Revised." then set end of resultCandidates to contents of candidate
                    end if
                end repeat
                if (count of resultCandidates) is 1 then
                    set comparisonDocument to item 1 of resultCandidates
                else
                    error "无法唯一识别 Word 的比较结果；请勿在比较过程中操作其他 Word 文档。"
                end if
                save as comparisonDocument file name destinationFile file format format document default
                close comparisonDocument saving no
                if sourceDocument is not comparisonDocument then
                    try
                        close sourceDocument saving no
                    end try
                end if
            on error errorMessage number errorNumber
                -- Avoid closing the user's unrelated documents or quitting Word.
                if sourceDocument is not missing value then
                    try
                        close sourceDocument saving no
                    end try
                end if
                error errorMessage number errorNumber
            end try
        end tell
    end timeout
    return item 3 of arguments
end run
