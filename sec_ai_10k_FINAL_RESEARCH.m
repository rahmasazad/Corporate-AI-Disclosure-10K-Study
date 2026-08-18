%% sec_ai_10k_batch_index_v2.m
% MULTI-COMPANY SEC 10-K AI DISCLOSURE ANALYZER v2
% MATLAB R2025b-compatible; no Text Analytics Toolbox required.
%
% WHAT THIS SCRIPT DOES
% 1. Reads a list of public companies.
% 2. Downloads each company's 10-K filings from SEC EDGAR.
% 3. Extracts AI-related paragraphs.
% 4. Scores language for risk, caution, promotion, and specificity.
% 5. Classifies AI discussion into research categories:
%       Cybersecurity, Privacy, IP/Copyright, Bias/Fairness,
%       Regulation, Governance/Safety, Competition/Business,
%       Workforce, Accuracy/Reliability.
% 6. Compares each year with the prior year.
% 7. Calculates a transparent, researcher-defined AI Disclosure Index (0-100).
% 8. Combines all companies into one master Excel workbook.
%
% IMPORTANT RESEARCH NOTE
% The AI Disclosure Index is a descriptive research measure created for this
% project. It is NOT an SEC standard and does NOT determine whether a filing
% is truthful, misleading, adequate, or legally compliant.
%
% BEFORE RUNNING
% Replace contactEmail below with your real email address.

clear;
clc;
close all;

%% ======================== 1. USER SETTINGS ============================

contactEmail = "rahmasazad@gmail.com";

startFiscalYear = 2020;
endFiscalYear   = 2025;

requestPauseSeconds = 0.25;

% Similarity thresholds used for year-to-year comparison.
boilerplateThreshold = 0.80;
newLanguageThreshold = 0.35;

% Master output folder.
masterOutputFolder = fullfile(pwd, "SEC_AI_Master_Study");

% ----------------------------------------------------------------------
% COMPANY LIST
%
% Group should be one of:
%   "Flagged"  = company was actually subject to the SEC event you study
%   "Control"  = comparable company without that event
%   "Test"     = used only to test the program
%
% EventDate:
%   Use datetime(YYYY,MM,DD) for a real event.
%   Use NaT for controls/test companies.
%
% Microsoft and C3.ai are included only as TEST examples here.
% Do not call them SEC AI-enforcement cases without separate evidence.
% ----------------------------------------------------------------------

companies = table( ...
    ["Presto Automation Inc."; "PAR Technology Corporation"; "GitLab Inc."; "JFrog Ltd."; "Apple Inc."; "Alphabet Inc."], ...
    ["PRST"; "PAR"; "GTLB"; "FROG"; "AAPL"; "GOOGL"], ...
    ["1822145"; "708821"; "1653482"; "1800667"; "320193"; "1652044"], ...
    [datetime(2025,1,14); datetime(2025,1,14); datetime(2024,9,4); datetime(2024,9,4); datetime(2025,3,19); datetime(2025,3,19)], ...
    ["SEC Enforcement"; "Control"; "Securities Litigation"; "Control"; "Consumer Litigation"; "Control"], ...
    ["Presto-PAR"; "Presto-PAR"; "GitLab-JFrog"; "GitLab-JFrog"; "Apple-Alphabet"; "Apple-Alphabet"], ...
    [2023; 2020; 2021; 2020; 2020; 2020], ...
    'VariableNames', {'Company','Ticker','CIK','EventDate','Group','PairID','MinFiscalYear'});

%% ======================== 2. VALIDATION ================================

if contactEmail == "REPLACE_WITH_YOUR_EMAIL@example.com"
    error("Replace contactEmail near the top with your real email address.");
end

if ~isfolder(masterOutputFolder)
    mkdir(masterOutputFolder);
end

fprintf("\nMULTI-COMPANY SEC 10-K AI DISCLOSURE STUDY\n");
fprintf("Companies configured: %d\n", height(companies));
fprintf("Fiscal years: %d-%d\n\n", startFiscalYear, endFiscalYear);

%% ======================== 3. DICTIONARIES ==============================

dict.AI = [
    "\bartificial intelligence\b"
    "\bgenerative ai\b"
    "\bgenai\b"
    "\bmachine learning\b"
    "\bdeep learning\b"
    "\blarge language models?\b"
    "\bllms?\b"
    "\bneural networks?\b"
    "\bnatural language processing\b"
    "\bcomputer vision\b"
    "\bai[- ](?:based|enabled|driven|powered|related|generated)\b"
    "\bai systems?\b"
    "\bai models?\b"
    "\bai tools?\b"
    "\bai technologies\b"
    "\bai solutions\b"
    "\bai capabilities\b"
    "(?<![A-Za-z])AI(?![A-Za-z])"
];

dict.Risk = [
    "\brisk(?:s)?\b"
    "\bharm(?:s|ful)?\b"
    "\berror(?:s)?\b"
    "\bfail(?:s|ed|ure|ures|ing)?\b"
    "\binaccura(?:te|cy|cies)\b"
    "\bhallucinat(?:e|es|ed|ion|ions)\b"
    "\bbias(?:ed|es)?\b"
    "\bdiscriminat(?:e|ion|ory)\b"
    "\bprivacy\b"
    "\bsecurity\b"
    "\bcybersecurity\b"
    "\bliabilit(?:y|ies)\b"
    "\blitigation\b"
    "\bregulat(?:e|ed|ion|ions|ory)\b"
    "\bcompliance\b"
    "\bintellectual property\b"
    "\bcopyright\b"
    "\bpatent(?:s)?\b"
    "\btrade secret(?:s)?\b"
    "\breputational\b"
    "\bmisuse\b"
    "\babuse\b"
    "\bunauthorized\b"
    "\bethical\b"
    "\badverse(?:ly)?\b"
    "\bdisruption\b"
];

dict.Caution = [
    "\bmay\b"
    "\bmight\b"
    "\bcould\b"
    "\bpotential(?:ly)?\b"
    "\bpossibly\b"
    "\buncertain(?:ty|ties)?\b"
    "\bcannot guarantee\b"
    "\bno assurance\b"
    "\bsubject to\b"
    "\bdepend(?:s|ed|ing)? on\b"
    "\bmay not\b"
    "\bcould adversely\b"
    "\bmay adversely\b"
    "\bmay be unable\b"
    "\bwe cannot\b"
];

dict.Promotion = [
    "\bleading\b"
    "\btransformative\b"
    "\brevolutionary\b"
    "\bcutting[- ]edge\b"
    "\binnovative\b"
    "\bbest[- ]in[- ]class\b"
    "\bstate[- ]of[- ]the[- ]art\b"
    "\bworld[- ]class\b"
    "\badvanced\b"
    "\bpowerful\b"
    "\bnext[- ]generation\b"
    "\bindustry[- ]leading\b"
    "\bmarket[- ]leading\b"
    "\bgroundbreaking\b"
    "\bunmatched\b"
    "\bdifferentiated\b"
];

dict.Specificity = [
    "\b\d+(?:\.\d+)?%\b"
    "\$\s?\d+(?:\.\d+)?\s?(?:million|billion|thousand)?\b"
    "\b20\d{2}\b"
    "\b(?:gdpr|ccpa|eu ai act|nist)\b"
    "\b(?:copilot|azure|openai|gemini|watson|firefly|einstein)\b"
];

% Research categories.
categories = struct();

categories.Cybersecurity = [
    "\bcybersecurity\b"
    "\bcyber attack(?:s)?\b"
    "\bsecurity breach(?:es)?\b"
    "\bdata breach(?:es)?\b"
    "\bmalware\b"
    "\bransomware\b"
    "\bthreat actor(?:s)?\b"
    "\bunauthorized access\b"
];

categories.Privacy = [
    "\bprivacy\b"
    "\bpersonal data\b"
    "\bpersonal information\b"
    "\bdata protection\b"
    "\bdata collection\b"
    "\bconsent\b"
    "\bgdpr\b"
    "\bccpa\b"
];

categories.IPCopyright = [
    "\bintellectual property\b"
    "\bcopyright\b"
    "\bpatent(?:s)?\b"
    "\btrade secret(?:s)?\b"
    "\blicens(?:e|es|ed|ing)\b"
    "\binfring(?:e|ement|ements)\b"
    "\btraining data\b"
];

categories.BiasFairness = [
    "\bbias(?:ed|es)?\b"
    "\bfairness\b"
    "\bdiscriminat(?:e|ion|ory)\b"
    "\bequity\b"
    "\bresponsible ai\b"
    "\bethical ai\b"
];

categories.Regulation = [
    "\bregulat(?:e|ed|ion|ions|ory)\b"
    "\blegislat(?:e|ed|ion|ive)\b"
    "\blaw(?:s)?\b"
    "\bcompliance\b"
    "\beu ai act\b"
    "\bnist\b"
    "\bgovernment oversight\b"
];

categories.GovernanceSafety = [
    "\bgovernance\b"
    "\boversight\b"
    "\bsafety\b"
    "\bresponsible ai\b"
    "\bethical\b"
    "\brisk management\b"
    "\binternal control(?:s)?\b"
    "\bmonitor(?:ing|ed)?\b"
    "\baudit(?:s|ed|ing)?\b"
];

categories.CompetitionBusiness = [
    "\bcompetition\b"
    "\bcompetitive\b"
    "\bcompetitor(?:s)?\b"
    "\bmarket share\b"
    "\bbusiness model\b"
    "\bcommercialization\b"
    "\bproductivity\b"
    "\bcustomer demand\b"
    "\brevenue\b"
    "\bprofit(?:s|ability)?\b"
];

categories.Workforce = [
    "\bworkforce\b"
    "\bemployee(?:s)?\b"
    "\bjob(?:s)?\b"
    "\blabor\b"
    "\bautomation\b"
    "\bdisplacement\b"
    "\btraining\b"
    "\breskilling\b"
    "\bhiring\b"
];

categories.AccuracyReliability = [
    "\baccuracy\b"
    "\binaccura(?:te|cy|cies)\b"
    "\breliability\b"
    "\bunreliable\b"
    "\berror(?:s)?\b"
    "\bhallucinat(?:e|es|ed|ion|ions)\b"
    "\bfalse output(?:s)?\b"
    "\bmisinformation\b"
    "\bquality\b"
];

categoryNames = string(fieldnames(categories));

%% ======================== 4. RUN ALL COMPANIES =========================

masterSummary = table();
masterParagraphs = table();
masterComparisons = table();
masterCategories = table();

for companyIndex = 1:height(companies)

    companyRow = companies(companyIndex,:);

    fprintf("\n====================================================\n");
    fprintf("ANALYZING %s (%s)\n", ...
        companyRow.Company, companyRow.Ticker);
    fprintf("====================================================\n");

    [summaryRows, paragraphRows, comparisonRows, categoryRows] = ...
        analyzeOneCompany( ...
        companyRow, ...
        startFiscalYear, ...
        endFiscalYear, ...
        contactEmail, ...
        requestPauseSeconds, ...
        boilerplateThreshold, ...
        newLanguageThreshold, ...
        masterOutputFolder, ...
        dict, ...
        categories, ...
        categoryNames);

    masterSummary = [masterSummary; summaryRows]; %#ok<AGROW>
    masterParagraphs = [masterParagraphs; paragraphRows]; %#ok<AGROW>
    masterComparisons = [masterComparisons; comparisonRows]; %#ok<AGROW>
    masterCategories = [masterCategories; categoryRows]; %#ok<AGROW>
end

%% ======================== 5. CASE-LEVEL BEFORE/AFTER =================

caseChanges = buildCaseChanges(masterSummary);
pairDifferences = buildPairDifferences(caseChanges);

%% ======================== 6. MASTER EXPORT =============================

timestamp = string(datetime("now","Format","yyyyMMdd_HHmmss"));
masterWorkbook = fullfile(masterOutputFolder, ...
    "MASTER_AI_DISCLOSURE_STUDY_" + timestamp + ".xlsx");

masterSummaryCSV = fullfile(masterOutputFolder, ...
    "MASTER_Company_Year_Summary.csv");

masterCategoryCSV = fullfile(masterOutputFolder, ...
    "MASTER_Category_Results.csv");

masterComparisonCSV = fullfile(masterOutputFolder, ...
    "MASTER_Year_Comparisons.csv");

writetable(masterSummary, masterWorkbook, "Sheet", "Company-Year Summary");
writetable(masterCategories, masterWorkbook, "Sheet", "Category Results");
writetable(masterComparisons, masterWorkbook, "Sheet", "Year Comparisons");
writetable(masterParagraphs, masterWorkbook, "Sheet", "AI Paragraphs");
writetable(caseChanges, masterWorkbook, "Sheet", "Case Changes");
writetable(pairDifferences, masterWorkbook, "Sheet", "Pair Differences");

writetable(masterSummary, masterSummaryCSV);
writetable(masterCategories, masterCategoryCSV);
writetable(masterComparisons, masterComparisonCSV);

save(fullfile(masterOutputFolder, "MASTER_AI_STUDY.mat"), ...
    "masterSummary", "masterParagraphs", ...
    "masterComparisons", "masterCategories", "caseChanges", ...
    "pairDifferences", "companies");

makeMasterCharts(masterSummary, masterOutputFolder);

fprintf("\n====================================================\n");
fprintf("MASTER STUDY COMPLETE\n");
fprintf("Open this workbook first:\n%s\n", masterWorkbook);
fprintf("====================================================\n\n");

displayVars = [
    "Company"
    "Group"
    "FiscalYear"
    "AIParagraphCount"
    "AIMentionsPer10000Words"
    "RiskDiversityCount"
    "NewLanguageShare"
    "AIDisclosureIndex"
];

disp(masterSummary(:, displayVars));

%% ======================== LOCAL FUNCTIONS ==============================

function [summaryRows, allParagraphRows, comparisonRows, categoryRows] = ...
    analyzeOneCompany( ...
    companyRow, startFiscalYear, endFiscalYear, contactEmail, ...
    requestPauseSeconds, boilerplateThreshold, newLanguageThreshold, ...
    masterOutputFolder, dict, categories, categoryNames)

    companyName = companyRow.Company;
    ticker = companyRow.Ticker;
    cik = companyRow.CIK;
    eventDate = companyRow.EventDate;
    group = companyRow.Group;
    pairID = companyRow.PairID;
    minFiscalYear = companyRow.MinFiscalYear;

    companyFolder = fullfile(masterOutputFolder, ticker);

    if ~isfolder(companyFolder)
        mkdir(companyFolder);
    end

    cik10 = pad(string(cik), 10, "left", "0");
    submissionsURL = ...
        "https://data.sec.gov/submissions/CIK" + cik10 + ".json";

    jsonOptions = makeSECWebOptions(contactEmail, "json");

    submissions = webread(submissionsURL, jsonOptions);
    pause(requestPauseSeconds);

    filings = buildFilingsTable(submissions, contactEmail, requestPauseSeconds);

    effectiveStartYear = max(startFiscalYear, minFiscalYear);

    keep = filings.Form == "10-K" & ...
           filings.FiscalYear >= effectiveStartYear & ...
           filings.FiscalYear <= endFiscalYear;

    selected = filings(keep,:);
    selected = sortrows(selected, ["FiscalYear","FilingDate"]);

    if isempty(selected)
        warning("No 10-K filings found for %s.", ticker);
        summaryRows = table();
        allParagraphRows = table();
        comparisonRows = table();
        categoryRows = table();
        return;
    end

    [~, ia] = unique(selected.FiscalYear, "stable");
    selected = selected(ia,:);
    selected = sortrows(selected, "FiscalYear");

    summaryRows = table();
    allParagraphRows = table();
    categoryRows = table();

    for r = 1:height(selected)

        fiscalYear = selected.FiscalYear(r);
        accession = selected.AccessionNumber(r);
        primaryDocument = selected.PrimaryDocument(r);
        filingDate = selected.FilingDate(r);
        reportDate = selected.ReportDate(r);

        accessionNoDashes = erase(accession, "-");

        filingURL = "https://www.sec.gov/Archives/edgar/data/" + ...
            string(str2double(cik)) + "/" + ...
            accessionNoDashes + "/" + primaryDocument;

        fprintf("Processing %s fiscal year %d...\n", ticker, fiscalYear);

        htmlFile = fullfile(companyFolder, ...
            sprintf("%s_%d_10K.html", ticker, fiscalYear));

        textFile = fullfile(companyFolder, ...
            sprintf("%s_%d_10K_text.txt", ticker, fiscalYear));

        if isfile(htmlFile)
            htmlCode = string(fileread(htmlFile));
        else
            textOptions = makeSECWebOptions(contactEmail, "text");
            htmlCode = string(webread(filingURL, textOptions));
            writeTextFile(htmlFile, htmlCode);
            pause(requestPauseSeconds);
        end

        filingText = extractFilingText(htmlCode);
        writeTextFile(textFile, filingText);

        totalWords = countWords(filingText);
        paragraphs = splitIntoParagraphs(filingText);

        aiMask = matchesAnyPattern(paragraphs, dict.AI);
        aiParagraphs = paragraphs(aiMask);

        aiMentionCount = countPatternMatches(filingText, dict.AI);
        aiParagraphCount = numel(aiParagraphs);
        aiMentionsPer10000 = ...
            safeDivide(aiMentionCount, totalWords) * 10000;

        fprintf("  %d AI paragraphs; %d AI matches.\n", ...
            aiParagraphCount, aiMentionCount);

        if aiParagraphCount == 0
            scored = emptyScoreTable(categoryNames);
        else
            scored = scoreParagraphs( ...
                aiParagraphs, dict, categories, categoryNames);
        end

        relativePeriod = classifyRelativePeriod(filingDate, eventDate);

        categoryCounts = zeros(1, numel(categoryNames));
        categoryShares = zeros(1, numel(categoryNames));

        for c = 1:numel(categoryNames)
            countVariable = categoryNames(c) + "Count";
            categoryCounts(c) = sum(scored.(countVariable));
            categoryShares(c) = safeDivide( ...
                sum(scored.(countVariable) > 0), aiParagraphCount);
        end

        riskDiversityCount = sum(categoryCounts > 0);

        summaryOne = table( ...
            companyName, ticker, cik, group, pairID, fiscalYear, ...
            reportDate, filingDate, eventDate, relativePeriod, ...
            filingURL, totalWords, aiParagraphCount, ...
            aiMentionCount, aiMentionsPer10000, ...
            sum(scored.RiskCount), ...
            sum(scored.CautionCount), ...
            sum(scored.PromotionCount), ...
            sum(scored.SpecificityCount), ...
            safeDivide(sum(scored.RiskCount > 0), aiParagraphCount), ...
            safeDivide(sum(scored.CautionCount > 0), aiParagraphCount), ...
            safeDivide(sum(scored.PromotionCount > 0), aiParagraphCount), ...
            safeDivide(sum(scored.SpecificityCount > 0), aiParagraphCount), ...
            riskDiversityCount, ...
            NaN, ... % NewLanguageShare filled after comparison.
            NaN, ... % BoilerplateShare filled after comparison.
            NaN, ... % AI Disclosure Index filled later.
            'VariableNames', { ...
            'Company','Ticker','CIK','Group','PairID','FiscalYear', ...
            'ReportDate','FilingDate','EventDate','RelativePeriod', ...
            'FilingURL','TotalWords','AIParagraphCount', ...
            'AIMentionCount','AIMentionsPer10000Words', ...
            'RiskTermCount','CautionTermCount','PromotionTermCount', ...
            'SpecificityIndicatorCount','RiskParagraphShare', ...
            'CautionParagraphShare','PromotionParagraphShare', ...
            'SpecificParagraphShare','RiskDiversityCount', ...
            'NewLanguageShare','BoilerplateShare', ...
            'AIDisclosureIndex'});

        summaryRows = [summaryRows; summaryOne]; %#ok<AGROW>

        for c = 1:numel(categoryNames)
            categoryOne = table( ...
                companyName, ticker, group, pairID, fiscalYear, ...
                categoryNames(c), categoryCounts(c), ...
                categoryShares(c), ...
                'VariableNames', { ...
                'Company','Ticker','Group','PairID','FiscalYear', ...
                'Category','TermCount','ParagraphShare'});

            categoryRows = [categoryRows; categoryOne]; %#ok<AGROW>
        end

        if aiParagraphCount > 0
            metadata = table( ...
                repmat(companyName, aiParagraphCount, 1), ...
                repmat(ticker, aiParagraphCount, 1), ...
                repmat(cik, aiParagraphCount, 1), ...
                repmat(group, aiParagraphCount, 1), ...
                repmat(pairID, aiParagraphCount, 1), ...
                repmat(fiscalYear, aiParagraphCount, 1), ...
                repmat(reportDate, aiParagraphCount, 1), ...
                repmat(filingDate, aiParagraphCount, 1), ...
                repmat(relativePeriod, aiParagraphCount, 1), ...
                repmat(filingURL, aiParagraphCount, 1), ...
                (1:aiParagraphCount)', ...
                'VariableNames', { ...
                'Company','Ticker','CIK','Group','PairID','FiscalYear', ...
                'ReportDate','FilingDate','RelativePeriod', ...
                'FilingURL','ParagraphNumber'});

            allParagraphRows = [allParagraphRows; metadata scored]; %#ok<AGROW>
        end
    end

    summaryRows = sortrows(summaryRows, "FiscalYear");

    if ~isempty(allParagraphRows)
        allParagraphRows = sortrows(allParagraphRows, ...
            ["FiscalYear","ParagraphNumber"]);
    end

    comparisonRows = table();

    yearsFound = summaryRows.FiscalYear;

    for k = 2:numel(yearsFound)

        priorYear = yearsFound(k-1);
        currentYear = yearsFound(k);

        priorParagraphs = strings(0,1);
        currentParagraphs = strings(0,1);

        if ~isempty(allParagraphRows)
            priorParagraphs = allParagraphRows.Paragraph( ...
                allParagraphRows.FiscalYear == priorYear);

            currentParagraphs = allParagraphRows.Paragraph( ...
                allParagraphRows.FiscalYear == currentYear);
        end

        [details, stats] = compareParagraphSets( ...
            priorParagraphs, currentParagraphs, ...
            boilerplateThreshold, newLanguageThreshold);

        if ~isempty(details)
            detailFile = fullfile(companyFolder, ...
                sprintf("%s_%d_vs_%d_AI_Paragraphs.csv", ...
                ticker, priorYear, currentYear));

            writetable(details, detailFile);
        end

        comparisonOne = table( ...
            companyName, ticker, group, pairID, priorYear, currentYear, ...
            stats.MeanMaxSimilarity, ...
            stats.MedianMaxSimilarity, ...
            stats.BoilerplateCount, ...
            stats.BoilerplateShare, ...
            stats.NewLanguageCount, ...
            stats.NewLanguageShare, ...
            'VariableNames', { ...
            'Company','Ticker','Group','PairID','PriorFiscalYear', ...
            'CurrentFiscalYear','MeanMaximumSimilarity', ...
            'MedianMaximumSimilarity','BoilerplateParagraphCount', ...
            'BoilerplateParagraphShare','NewLanguageParagraphCount', ...
            'NewLanguageParagraphShare'});

        comparisonRows = [comparisonRows; comparisonOne]; %#ok<AGROW>

        currentIndex = find(summaryRows.FiscalYear == currentYear, 1);

        summaryRows.NewLanguageShare(currentIndex) = ...
            stats.NewLanguageShare;

        summaryRows.BoilerplateShare(currentIndex) = ...
            stats.BoilerplateShare;
    end

    % First year has no prior filing for comparison.
    if height(summaryRows) >= 1
        summaryRows.NewLanguageShare(1) = 0;
        summaryRows.BoilerplateShare(1) = 0;
    end

    % Calculate the transparent 0-100 index.
    for r = 1:height(summaryRows)

        fiscalYear = summaryRows.FiscalYear(r);

        yearCategories = categoryRows( ...
            categoryRows.FiscalYear == fiscalYear, :);

        regulatoryShare = getCategoryShare( ...
            yearCategories, "Regulation");

        governanceShare = getCategoryShare( ...
            yearCategories, "GovernanceSafety");

        frequencyComponent = min(20, ...
            summaryRows.AIMentionsPer10000Words(r));

        noveltyComponent = ...
            20 * summaryRows.NewLanguageShare(r);

        diversityComponent = ...
            20 * safeDivide(summaryRows.RiskDiversityCount(r), ...
            numel(categoryNames));

        regulatoryComponent = ...
            15 * min(1, regulatoryShare);

        specificityComponent = ...
            15 * min(1, summaryRows.SpecificParagraphShare(r));

        governanceComponent = ...
            10 * min(1, governanceShare);

        indexScore = frequencyComponent + ...
            noveltyComponent + ...
            diversityComponent + ...
            regulatoryComponent + ...
            specificityComponent + ...
            governanceComponent;

        summaryRows.AIDisclosureIndex(r) = round(indexScore, 2);
    end

    companyWorkbook = fullfile(companyFolder, ...
        ticker + "_AI_Analysis.xlsx");

    writetable(summaryRows, companyWorkbook, "Sheet", "Summary");
    writetable(categoryRows, companyWorkbook, "Sheet", "Categories");
    writetable(comparisonRows, companyWorkbook, "Sheet", "Comparisons");
    writetable(allParagraphRows, companyWorkbook, "Sheet", "AI Paragraphs");

    makeCompanyCharts(summaryRows, categoryRows, ...
        companyFolder, ticker);
end

function caseChanges = buildCaseChanges(masterSummary)

    caseChanges = table();
    tickers = unique(masterSummary.Ticker, "stable");

    for i = 1:numel(tickers)
        rows = masterSummary(masterSummary.Ticker == tickers(i), :);
        rows = sortrows(rows, "FilingDate");

        eventDate = rows.EventDate(1);
        if isnat(eventDate)
            continue;
        end

        before = rows(rows.FilingDate < eventDate, :);
        after  = rows(rows.FilingDate >= eventDate, :);

        hasBefore = ~isempty(before);
        hasAfter = ~isempty(after);

        if hasBefore
            beforeRow = before(end,:);
            beforeFY = beforeRow.FiscalYear;
            beforeIndex = beforeRow.AIDisclosureIndex;
            beforeMentions = beforeRow.AIMentionCount;
            beforeParagraphs = beforeRow.AIParagraphCount;
            beforeRiskShare = beforeRow.RiskParagraphShare;
            beforeSpecificShare = beforeRow.SpecificParagraphShare;
        else
            beforeFY = NaN; beforeIndex = NaN; beforeMentions = NaN;
            beforeParagraphs = NaN; beforeRiskShare = NaN; beforeSpecificShare = NaN;
        end

        if hasAfter
            afterRow = after(1,:);
            afterFY = afterRow.FiscalYear;
            afterIndex = afterRow.AIDisclosureIndex;
            afterMentions = afterRow.AIMentionCount;
            afterParagraphs = afterRow.AIParagraphCount;
            afterRiskShare = afterRow.RiskParagraphShare;
            afterSpecificShare = afterRow.SpecificParagraphShare;
        else
            afterFY = NaN; afterIndex = NaN; afterMentions = NaN;
            afterParagraphs = NaN; afterRiskShare = NaN; afterSpecificShare = NaN;
        end

        one = table( ...
            rows.Company(1), rows.Ticker(1), rows.Group(1), rows.PairID(1), ...
            eventDate, hasBefore, hasAfter, ...
            beforeFY, afterFY, beforeIndex, afterIndex, afterIndex-beforeIndex, ...
            beforeMentions, afterMentions, afterMentions-beforeMentions, ...
            beforeParagraphs, afterParagraphs, afterParagraphs-beforeParagraphs, ...
            beforeRiskShare, afterRiskShare, afterRiskShare-beforeRiskShare, ...
            beforeSpecificShare, afterSpecificShare, afterSpecificShare-beforeSpecificShare, ...
            'VariableNames', { ...
            'Company','Ticker','Group','PairID','EventDate','HasPreEvent10K','HasPostEvent10K', ...
            'PreFiscalYear','PostFiscalYear','PreIndex','PostIndex','IndexChange', ...
            'PreAIMentions','PostAIMentions','AIMentionChange', ...
            'PreAIParagraphs','PostAIParagraphs','AIParagraphChange', ...
            'PreRiskShare','PostRiskShare','RiskShareChange', ...
            'PreSpecificShare','PostSpecificShare','SpecificShareChange'});

        caseChanges = [caseChanges; one]; %#ok<AGROW>
    end
end

function pairDifferences = buildPairDifferences(caseChanges)

    pairDifferences = table();
    if isempty(caseChanges)
        return;
    end

    pairs = unique(caseChanges.PairID, "stable");

    for i = 1:numel(pairs)
        rows = caseChanges(caseChanges.PairID == pairs(i), :);
        treated = rows(rows.Group ~= "Control", :);
        control = rows(rows.Group == "Control", :);

        if height(treated) ~= 1 || height(control) ~= 1
            continue;
        end

        usable = treated.HasPreEvent10K && treated.HasPostEvent10K && ...
                 control.HasPreEvent10K && control.HasPostEvent10K;

        one = table( ...
            pairs(i), treated.Company(1), control.Company(1), treated.Group(1), usable, ...
            treated.IndexChange, control.IndexChange, ...
            treated.IndexChange-control.IndexChange, ...
            treated.AIMentionChange, control.AIMentionChange, ...
            treated.AIMentionChange-control.AIMentionChange, ...
            treated.RiskShareChange, control.RiskShareChange, ...
            treated.RiskShareChange-control.RiskShareChange, ...
            treated.SpecificShareChange, control.SpecificShareChange, ...
            treated.SpecificShareChange-control.SpecificShareChange, ...
            'VariableNames', { ...
            'PairID','CaseCompany','ControlCompany','EventType','UsableBeforeAfterPair', ...
            'CaseIndexChange','ControlIndexChange','DifferenceInIndexChange', ...
            'CaseAIMentionChange','ControlAIMentionChange','DifferenceInAIMentionChange', ...
            'CaseRiskShareChange','ControlRiskShareChange','DifferenceInRiskShareChange', ...
            'CaseSpecificShareChange','ControlSpecificShareChange','DifferenceInSpecificShareChange'});

        pairDifferences = [pairDifferences; one]; %#ok<AGROW>
    end
end

function value = getCategoryShare(categoryRows, categoryName)

    match = categoryRows.Category == categoryName;

    if any(match)
        value = categoryRows.ParagraphShare(find(match,1));
    else
        value = 0;
    end
end

function options = makeSECWebOptions(contactEmail, contentType)

    userAgentText = ...
        "Student academic research " + string(contactEmail);

    headers = [
        "Accept-Encoding", "gzip, deflate"
    ];

    if strcmpi(string(contentType), "json")
        options = weboptions( ...
            "ContentType", "json", ...
            "Timeout", 60, ...
            "UserAgent", userAgentText, ...
            "HeaderFields", headers);
    else
        options = weboptions( ...
            "ContentType", "text", ...
            "Timeout", 90, ...
            "UserAgent", userAgentText, ...
            "HeaderFields", headers);
    end
end

function filings = buildFilingsTable(submissions, contactEmail, requestPauseSeconds)

    % The SEC's main company-submissions JSON contains a "recent" block.
    % Older filings can be stored in separate JSON files named in
    % submissions.filings.files. Read both so long filing histories are not
    % silently truncated.

    recentTable = filingStructToTable(submissions.filings.recent);
    filings = recentTable;

    if isfield(submissions.filings, "files") && ~isempty(submissions.filings.files)
        archiveFiles = submissions.filings.files;

        if isstruct(archiveFiles)
            for i = 1:numel(archiveFiles)
                if ~isfield(archiveFiles(i), "name")
                    continue;
                end

                archiveName = string(archiveFiles(i).name);
                archiveURL = "https://data.sec.gov/submissions/" + archiveName;

                try
                    archiveOptions = makeSECWebOptions(contactEmail, "json");
                    archiveData = webread(archiveURL, archiveOptions);
                    pause(requestPauseSeconds);

                    archiveTable = filingStructToTable(archiveData);
                    filings = [filings; archiveTable]; %#ok<AGROW>
                catch ME
                    warning("Could not read SEC archive %s: %s", ...
                        archiveName, ME.message);
                end
            end
        end
    end

    % Remove duplicates that may appear across recent/archive boundaries.
    if ~isempty(filings)
        [~, ia] = unique(filings.AccessionNumber, "stable");
        filings = filings(ia,:);
        filings = sortrows(filings, "FilingDate");
    end
end

function filings = filingStructToTable(r)

    filings = table( ...
        string(r.form(:)), ...
        datetime(string(r.filingDate(:)), "InputFormat", "yyyy-MM-dd"), ...
        datetime(string(r.reportDate(:)), "InputFormat", "yyyy-MM-dd"), ...
        string(r.accessionNumber(:)), ...
        string(r.primaryDocument(:)), ...
        'VariableNames', { ...
        'Form','FilingDate','ReportDate', ...
        'AccessionNumber','PrimaryDocument'});

    filings.FiscalYear = year(filings.ReportDate);
end

function text = extractFilingText(htmlCode)

    htmlChar = char(string(htmlCode));

    htmlChar = regexprep(htmlChar, ...
        '(?is)<script\b[^>]*>.*?</script>', ' ');

    htmlChar = regexprep(htmlChar, ...
        '(?is)<style\b[^>]*>.*?</style>', ' ');

    htmlChar = regexprep(htmlChar, ...
        '(?is)<ix:header\b[^>]*>.*?</ix:header>', ' ');

    htmlChar = regexprep(htmlChar, ...
        ['(?i)</?(?:p|div|section|article|li|tr|td|th|' ...
        'h1|h2|h3|h4|h5|h6|br)[^>]*>'], ...
        sprintf('\n'));

    htmlChar = regexprep(htmlChar, '(?is)<[^>]+>', ' ');

    htmlChar = strrep(htmlChar, '&nbsp;', ' ');
    htmlChar = strrep(htmlChar, '&#160;', ' ');
    htmlChar = strrep(htmlChar, '&amp;', '&');
    htmlChar = strrep(htmlChar, '&quot;', '"');
    htmlChar = strrep(htmlChar, '&#39;', char(39));
    htmlChar = strrep(htmlChar, '&lt;', '<');
    htmlChar = strrep(htmlChar, '&gt;', '>');

    htmlChar = regexprep(htmlChar, ...
        '&#x[0-9A-Fa-f]+;', ' ');

    htmlChar = regexprep(htmlChar, ...
        '&#[0-9]+;', ' ');

    htmlChar = regexprep(htmlChar, ...
        '\r\n?', sprintf('\n'));

    htmlChar = regexprep(htmlChar, ...
        '[ \t\f\v]+', ' ');

    htmlChar = regexprep(htmlChar, ...
        '\n[ \t]+', sprintf('\n'));

    htmlChar = regexprep(htmlChar, ...
        '[ \t]+\n', sprintf('\n'));

    htmlChar = regexprep(htmlChar, ...
        '\n{3,}', sprintf('\n\n'));

    text = strip(string(htmlChar));
end

function paragraphs = splitIntoParagraphs(text)

    raw = splitlines(string(text));
    raw = strip(raw);
    raw(raw == "") = [];
    raw = raw(strlength(raw) >= 40);

    paragraphs = strings(0,1);
    buffer = "";

    for i = 1:numel(raw)

        if buffer == ""
            buffer = raw(i);
        else
            buffer = buffer + " " + raw(i);
        end

        endsSentence = ~isempty(regexp( ...
            char(buffer), '[.!?]["'')\]]?$', 'once'));

        if strlength(buffer) >= 250 || endsSentence
            paragraphs(end+1,1) = strip(buffer); %#ok<AGROW>
            buffer = "";
        end
    end

    if buffer ~= ""
        paragraphs(end+1,1) = strip(buffer);
    end

    paragraphs = regexprep(paragraphs, "\s+", " ");
    paragraphs = strip(paragraphs);
    paragraphs(paragraphs == "") = [];
    paragraphs = paragraphs(strlength(paragraphs) <= 5000);
end

function scored = scoreParagraphs( ...
    paragraphs, dict, categories, categoryNames)

    n = numel(paragraphs);

    riskCount = zeros(n,1);
    cautionCount = zeros(n,1);
    promotionCount = zeros(n,1);
    specificityCount = zeros(n,1);

    for i = 1:n
        riskCount(i) = ...
            countPatternMatches(paragraphs(i), dict.Risk);

        cautionCount(i) = ...
            countPatternMatches(paragraphs(i), dict.Caution);

        promotionCount(i) = ...
            countPatternMatches(paragraphs(i), dict.Promotion);

        specificityCount(i) = ...
            countPatternMatches(paragraphs(i), dict.Specificity);
    end

    scored = table( ...
        paragraphs(:), riskCount, cautionCount, ...
        promotionCount, specificityCount, ...
        riskCount > 0, cautionCount > 0, ...
        promotionCount > 0, specificityCount > 0, ...
        'VariableNames', { ...
        'Paragraph','RiskCount','CautionCount', ...
        'PromotionCount','SpecificityCount', ...
        'FlaggedRisk','FlaggedCaution', ...
        'FlaggedPromotion','FlaggedSpecific'});

    for c = 1:numel(categoryNames)

        categoryName = categoryNames(c);
        countVariable = categoryName + "Count";
        flagVariable = "Flagged" + categoryName;

        categoryCount = zeros(n,1);

        patterns = categories.(categoryName);

        for i = 1:n
            categoryCount(i) = ...
                countPatternMatches(paragraphs(i), patterns);
        end

        scored.(countVariable) = categoryCount;
        scored.(flagVariable) = categoryCount > 0;
    end
end

function tbl = emptyScoreTable(categoryNames)

    tbl = table( ...
        strings(0,1), ...
        zeros(0,1), zeros(0,1), ...
        zeros(0,1), zeros(0,1), ...
        false(0,1), false(0,1), ...
        false(0,1), false(0,1), ...
        'VariableNames', { ...
        'Paragraph','RiskCount','CautionCount', ...
        'PromotionCount','SpecificityCount', ...
        'FlaggedRisk','FlaggedCaution', ...
        'FlaggedPromotion','FlaggedSpecific'});

    for c = 1:numel(categoryNames)
        tbl.(categoryNames(c) + "Count") = zeros(0,1);
        tbl.("Flagged" + categoryNames(c)) = false(0,1);
    end
end

function tf = matchesAnyPattern(textArray, patterns)

    tf = false(size(textArray));
    textCells = cellstr(textArray);

    for i = 1:numel(patterns)
        patternChar = repairRegexPattern(patterns(i));
        hits = regexpi(textCells, patternChar, "once");
        tf = tf | ~cellfun("isempty", hits);
    end
end

function total = countPatternMatches(text, patterns)

    total = 0;
    textChar = char(text);

    for i = 1:numel(patterns)
        patternChar = repairRegexPattern(patterns(i));

        total = total + numel(regexpi( ...
            textChar, patternChar, "match"));
    end
end

function patternChar = repairRegexPattern(patternValue)

    % MATLAB regexp does not treat \b as the same word-boundary token
    % used in many other regex engines. Remove both possible forms so
    % the intended keyword or phrase can still match.
    patternChar = char(patternValue);

    % Remove actual backspace characters.
    patternChar = strrep(patternChar, char(8), '');

    % Remove literal backslash-b text.
    patternChar = strrep(patternChar, '\b', '');
end

function n = countWords(text)

    words = regexp(char(text), ...
        '[A-Za-z][A-Za-z''-]*', "match");

    n = numel(words);
end

function result = safeDivide(numerator, denominator)

    if denominator == 0
        result = 0;
    else
        result = numerator / denominator;
    end
end

function label = classifyRelativePeriod(filingDate, eventDate)

    if isnat(eventDate)
        label = "No event specified";
    elseif filingDate < eventDate
        label = "Before event";
    else
        label = "After event";
    end
end

function [details, stats] = compareParagraphSets( ...
    priorParagraphs, currentParagraphs, ...
    boilerplateThreshold, newLanguageThreshold)

    stats = struct( ...
        "MeanMaxSimilarity", 0, ...
        "MedianMaxSimilarity", 0, ...
        "BoilerplateCount", 0, ...
        "BoilerplateShare", 0, ...
        "NewLanguageCount", 0, ...
        "NewLanguageShare", 0);

    if isempty(currentParagraphs)
        details = table();
        return;
    end

    nCurrent = numel(currentParagraphs);
    maxSimilarity = zeros(nCurrent,1);
    closestPriorParagraph = strings(nCurrent,1);

    if isempty(priorParagraphs)
        closestPriorParagraph(:) = "";
    else
        priorTokens = cellfun( ...
            @tokenSet, cellstr(priorParagraphs), ...
            "UniformOutput", false);

        for i = 1:nCurrent

            currentTokens = ...
                tokenSet(char(currentParagraphs(i)));

            similarities = ...
                zeros(numel(priorParagraphs),1);

            for j = 1:numel(priorParagraphs)
                similarities(j) = ...
                    jaccardSimilarity( ...
                    currentTokens, priorTokens{j});
            end

            [maxSimilarity(i), bestIndex] = ...
                max(similarities);

            closestPriorParagraph(i) = ...
                priorParagraphs(bestIndex);
        end
    end

    isBoilerplate = ...
        maxSimilarity >= boilerplateThreshold;

    isNewLanguage = ...
        maxSimilarity < newLanguageThreshold;

    details = table( ...
        currentParagraphs(:), ...
        closestPriorParagraph, ...
        maxSimilarity, ...
        isBoilerplate, ...
        isNewLanguage, ...
        'VariableNames', { ...
        'CurrentParagraph','ClosestPriorParagraph', ...
        'MaximumSimilarity','FlaggedBoilerplate', ...
        'FlaggedNewLanguage'});

    stats.MeanMaxSimilarity = mean(maxSimilarity);
    stats.MedianMaxSimilarity = median(maxSimilarity);
    stats.BoilerplateCount = sum(isBoilerplate);
    stats.BoilerplateShare = mean(isBoilerplate);
    stats.NewLanguageCount = sum(isNewLanguage);
    stats.NewLanguageShare = mean(isNewLanguage);
end

function tokens = tokenSet(text)

    words = regexp(lower(char(text)), ...
        '[a-z][a-z''-]+', "match");

    words = string(words);

    stopWords = [ ...
        "a","an","and","are","as","at","be","been","being", ...
        "but","by","for","from","had","has","have","he","her", ...
        "hers","him","his","i","in","into","is","it","its","of", ...
        "on","or","our","ours","she","that","the","their","theirs", ...
        "them","they","this","to","was","we","were","will","with", ...
        "you","your"];

    words = words(~ismember(words, stopWords));
    tokens = unique(words);
end

function value = jaccardSimilarity(tokensA, tokensB)

    if isempty(tokensA) && isempty(tokensB)
        value = 1;
        return;
    end

    unionCount = numel(union(tokensA, tokensB));

    if unionCount == 0
        value = 0;
    else
        value = numel(intersect(tokensA, tokensB)) ...
            / unionCount;
    end
end

function writeTextFile(filename, text)

    fileID = fopen(filename, "w", "n", "UTF-8");

    if fileID == -1
        error("Could not open file: %s", filename);
    end

    cleanupObject = onCleanup(@() fclose(fileID)); %#ok<NASGU>
    fprintf(fileID, "%s", char(text));
end

function makeCompanyCharts(summaryRows, categoryRows, ...
    companyFolder, ticker)

    years = summaryRows.FiscalYear;

    f1 = figure("Visible", "off");

    plot(years, summaryRows.AIDisclosureIndex, ...
        "-o", "LineWidth", 1.5);

    xlabel("Fiscal year");
    ylabel("AI Disclosure Index (0-100)");
    title(ticker + " AI Disclosure Index");
    ylim([0 100]);
    grid on;

    exportgraphics(f1, ...
        fullfile(companyFolder, ticker + "_AI_Index.png"), ...
        "Resolution", 200);

    close(f1);

    latestYear = max(categoryRows.FiscalYear);

    latest = categoryRows( ...
        categoryRows.FiscalYear == latestYear, :);

    f2 = figure("Visible", "off");

    bar(categorical(latest.Category), ...
        latest.ParagraphShare);

    xlabel("Category");
    ylabel("Share of AI paragraphs");
    title(ticker + " AI Categories, " + latestYear);
    ylim([0 1]);
    grid on;

    exportgraphics(f2, ...
        fullfile(companyFolder, ...
        ticker + "_Latest_Categories.png"), ...
        "Resolution", 200);

    close(f2);
end

function makeMasterCharts(masterSummary, masterOutputFolder)

    if isempty(masterSummary)
        return;
    end

    companies = unique(masterSummary.Ticker);

    f = figure("Visible", "off");
    hold on;

    for i = 1:numel(companies)

        rows = masterSummary.Ticker == companies(i);

        plot(masterSummary.FiscalYear(rows), ...
            masterSummary.AIDisclosureIndex(rows), ...
            "-o", "LineWidth", 1.5, ...
            "DisplayName", companies(i));
    end

    hold off;
    xlabel("Fiscal year");
    ylabel("AI Disclosure Index (0-100)");
    title("Company AI Disclosure Index Comparison");
    ylim([0 100]);
    legend("Location", "best");
    grid on;

    exportgraphics(f, ...
        fullfile(masterOutputFolder, ...
        "MASTER_AI_Index_Comparison.png"), ...
        "Resolution", 200);

    close(f);
end
