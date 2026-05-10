clc;
clear;

inputFile = 'AG2 U3 RTD (5mins).xlsx';   % <<< CHANGE THIS
outputPrefix = 'Repaired_';
timeStep = minutes(5);

[status, sheetNames] = xlsfinfo(inputFile);

for s = 1:length(sheetNames)
    sheet = sheetNames{s};
    fprintf('Processing %s...\n', sheet);

    % ======================================================
    % STEP 1: READ RAW DATA (No Auto-Conversion)
    % ======================================================
    rawParams = detectImportOptions(inputFile, 'Sheet', sheet);
    rawParams.DataRange = 'A:C'; 
    rawParams.VariableNamingRule = 'preserve';
    
    RawData = readcell(inputFile, 'Sheet', sheet);
    
    if ischar(RawData{1,1}) && isnan(str2double(RawData{1,1}))
         RawData(1,:) = [];
    end

    rawDates = RawData(:,1);
    rawPower = RawData(:,3); 

    % ======================================================
    % STEP 2: HYBRID DATE PARSING
    % ======================================================
    Time = NaT(height(rawDates), 1);
    
    for i = 1:length(rawDates)
        val = rawDates{i};
        if ischar(val) || isstring(val)
            try
                Time(i) = datetime(val, 'InputFormat', 'MM/dd/yyyy HH:mm');
            catch
                try
                    Time(i) = datetime(val);
                catch
                end
            end
        elseif isdatetime(val)
            Time(i) = val;
        end
    end
    
    Power = zeros(length(rawPower), 1);
    for i = 1:length(rawPower)
        if isnumeric(rawPower{i})
            Power(i) = rawPower{i};
        else
            Power(i) = str2double(rawPower{i});
        end
    end

    validIdx = ~isnat(Time) & ~isnan(Power);
    Time = Time(validIdx);
    Power = Power(validIdx);

    % ======================================================
    % STEP 3: THE "WRONG MONTH" FIX & PARSING SHEET NAME
    % ======================================================
    % FIX 1: Robust Year/Month Extraction (Handles "Sept" vs "Sep")
    yearStr = sheet(end-3:end);      % Last 4 chars = Year (2024)
    monthStr = sheet(1:end-4);       % Everything before = Month (Sept)
    
    yearNum = str2double(yearStr);
    
    % Handle "Sept" specifically as MATLAB usually expects "Sep"
    if strcmpi(monthStr, 'Sept')
        monthNum = 9;
    else
        monthNum = month(datetime(monthStr,'InputFormat','MMM'),'month');
    end
    
    % Detect Swap
    currentMonths = month(Time);
    wrongMonthIdx = (currentMonths ~= monthNum);
    
    if any(wrongMonthIdx)
        fprintf('  > Detected %d dates with swapped Day/Month. Fixing...\n', sum(wrongMonthIdx));
        badDates = Time(wrongMonthIdx);
        fixedDates = datetime(year(badDates), day(badDates), month(badDates), ...
                              hour(badDates), minute(badDates), second(badDates));
        Time(wrongMonthIdx) = fixedDates;
    end

    % ======================================================
    % STEP 4: STANDARD PROCESSING
    % ======================================================
    Time = dateshift(Time, 'start', 'minute');
    [Time, idx] = sort(Time);
    Power = Power(idx);

    [G, uniqueTime] = findgroups(Time);
    avgPower = splitapply(@mean, Power, G);
    TT = timetable(uniqueTime, avgPower, 'VariableNames', {'Power'});

    % Create perfect time grid
    startTime = datetime(yearNum, monthNum, 1, 0, 0, 0);
    endTime   = dateshift(startTime, 'end', 'month');
    endTime   = dateshift(endTime, 'end', 'day'); 
    fullTime  = (startTime:timeStep:endTime)';

    TTfixed = retime(TT, fullTime, 'fillwithmissing');

    outputFile = [outputPrefix sheet '.xlsx'];
    writetimetable(TTfixed, outputFile);
    
    fprintf('Finished %s (Rows: %d)\n\n', sheet, height(TTfixed));
end

disp('ALL MONTHS REPAIRED CORRECTLY');