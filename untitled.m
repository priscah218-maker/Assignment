clc;
clear;
close all;
while true
    fprintf('\n\n       MATLAB ASSIGNMENTS PROGRAM\n');
    fprintf('========================================\n');
    fprintf('1. Student Excel Assignment\n');
    fprintf('2. GPA/CGPA Calculator\n');
    fprintf('3. Leaf Image Processing\n');
    fprintf('4. Numerical Analysis\n');
    fprintf('5. Run ALL Assignments\n');
    fprintf('6. Exit\n');
    fprintf('========================================\n');
    choice = input('Enter your choice: ');
    switch choice
        case 1
            studentExcel();
        case 2
            GPA_CGPA();
        case 3
            leafProcessing();
        case 4
            numericalAnalysis();
        case 5
            fprintf('\n========== RUNNING ALL ASSIGNMENTS ==========\n');
            fprintf('\n--- Assignment 1 ---\n');
            studentExcel();
            fprintf('\n--- Assignment 2 ---\n');
            GPA_CGPA();
            fprintf('\n--- Assignment 3 ---\n');
            leafProcessing();
            fprintf('\n--- Assignment 4 ---\n');
            numericalAnalysis();
            fprintf('\n========================================\n');
            fprintf('     ALL ASSIGNMENTS COMPLETED\n');
        case 6
            fprintf('\nProgram ended.\n');
            break;
        otherwise
            fprintf('\nInvalid choice. Please try again.\n');
    end
end

%% ASSIGNMENT 1: STUDENT EXCEL
function studentExcel()
clc;
inputFile = 'student_details.xlsx';
students = readtable(inputFile,'VariableNamingRule', 'preserve');
disp('        STUDENT INFORMATION IMPORTED');
disp('================================================');
disp(students);

% DISPLAY THE NUMBER OF STUDENTS
numberOfStudents = height(students);
disp(' ');
disp(['Number of students: ', num2str(numberOfStudents)]);

% SELECT ALL THE REQUIRED STUDENT INFORMATION
selectedStudents = students;
disp(' ');
disp('        SELECTED STUDENT INFORMATION');
disp('================================================');
disp(selectedStudents);

% GROUP INFORMATION
groupName = 'Group 3';
groupMembers = {
    'WANGO LAWRENCE', 'BU/UP/2025/2147';
    'KABAGYENYI GIFT JOSEPHINE', 'BU/UP/2025/1539';
    'ACHAM PRISCA SOPHIE', 'BU/UP/2025/5143';
    'KUJO WINNIE', 'BU/X /2025/3266';
    'AHURRA ARTHUR ELLY', 'BU/UP/2025/2101'
    };
% NEW EXCEL FILE
outputFile = 'Group_3_Student_Information.xlsx';
writecell({'MATLAB STUDENT INFORMATION ASSIGNMENT'},outputFile,'Sheet', 1,'Range', 'A1');
writecell({' ', groupName},outputFile,'Sheet', 1,'Range', 'A2');
writecell({'GROUP MEMBERS'},outputFile,'Sheet', 1,'Range', 'A4');
writecell({'Name', 'Registration Number'},outputFile,'Sheet', 1,'Range', 'A5');
writecell(groupMembers,outputFile,'Sheet', 1,'Range', 'A6');
    % determine where the table starts
studentStartRow = 6 + size(groupMembers,1) + 2;
writecell({'STUDENT INFORMATION'},outputFile,'Sheet', 1,'Range', sprintf('A%d', studentStartRow));
writetable(selectedStudents,outputFile,'Sheet', 1,'Range', sprintf('A%d', studentStartRow + 1));
disp(' ');
disp(['Output file: ', outputFile]);

% CREATE A FOLDER FOR THE PLOTS
plotFolder = 'Student_Plots';
if ~exist(plotFolder, 'dir')
    mkdir(plotFolder);
end
    % genger graph
genderCounts = groupcounts(selectedStudents, 'GENDER');
figure;
bar(categorical(genderCounts.GENDER),genderCounts.GroupCount);
title('Number of Students by Gender');
xlabel('Gender');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Gender.png'));

    % age histogram
figure;
histogram(selectedStudents.AGE);
title('Age Distribution of Students');
xlabel('Age');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Age.png'));

    % tribe graph
tribeCounts = groupcounts(selectedStudents, 'TRIBE');
figure;
bar(categorical(tribeCounts.TRIBE),tribeCounts.GroupCount);
title('Number of Students by Tribe');
xlabel('Tribe');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Tribe.png'));

    % association grapha
associationCounts = groupcounts(selectedStudents, 'ASSOCIATION');
figure;
bar(categorical(associationCounts.ASSOCIATION),associationCounts.GroupCount);
title('Number of Students by Association');
xlabel('Association');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Association.png'));

    % hometown graph
hometownCounts = groupcounts(selectedStudents, 'HOMETOWN');
figure;
bar(categorical(hometownCounts.HOMETOWN),hometownCounts.GroupCount);
title('Number of Students by Hometown');
xlabel('Hometown');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Hometown.png'));

    % resident graph
residenceCounts = groupcounts(selectedStudents, 'RESIDENCE');
figure;
bar(categorical(residenceCounts.RESIDENCE),residenceCounts.GroupCount);
title('Number of Students by Residence');
xlabel('Residence');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Residence.png'));

    % religion graph
religionCounts = groupcounts(selectedStudents, 'RELIGION');
figure;
bar(categorical(religionCounts.RELIGION),religionCounts.GroupCount);
title('Number of Students by Religion');
xlabel('Religion');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Religion.png'));

    % hobby graph
hobbyCounts = groupcounts(selectedStudents, 'HOBBY');
figure;
bar(categorical(hobbyCounts.HOBBY),hobbyCounts.GroupCount);
title('Number of Students by Hobby');
xlabel('Hobby');
ylabel('Number of Students');
grid on;
exportgraphics(gcf,fullfile(plotFolder, 'Students_by_Hobby.png'));

disp(['Excel output file: ', outputFile]);
disp(['Plot folder: ', plotFolder]);
end

%% ASSIGNMENT 2: GPA / CGPA CALCULATOR
function GPA_CGPA()
clc;
format short;

fprintf('       BACHELOR OF WATER RESOURCES ENGINEERING\n');
fprintf('               GPA & CGPA CALCULATOR\n');
fprintf('============================================================\n\n');

% ENTER STUDENT DETAILS
studentName = input('Enter student name: ', 's');
regNo = input('Enter registration number: ', 's');
program = input('Enter programme of study: ', 's');
yearOfStudy = readNumber('Enter current year of study: ', 1, 4);
currentSemester = readNumber('Enter current semester: ', 1, 2);
fprintf('\n');

% CREATE THE PROGRAMME CURRICULUM
curriculum = cell(4,2);
% YEAR 1 - SEMESTER 1
curriculum{1,1} = {
    'WAR1101', 'Engineering Mathematics 1', 4;
    'WAR1102', 'Engineering Mechanics I', 4;
    'WAR1103', 'Circuits Theory', 4;
    'WAR1104', 'Computer Applications', 3;
    'WAR1105', 'Engineering Drawing', 4;
    'WAR1106', 'Communication skills', 3;
    'WAR1107', 'Sustainable Water Resources & Engineering Ethics', 3
    };
% YEAR 1 - SEMESTER 2
curriculum{1,2} = {
    'WAR1201', 'Engineering Mathematics II', 4;
    'WAR1202', 'Computer Aided Design', 3;
    'WAR1203', 'Thermodynamics', 4;
    'WAR1204', 'Fluid Mechanics', 4;
    'WAR1205', 'Environmental Science', 4;
    'WAR1206', 'Surveying for Engineers', 3
    };
% YEAR 2 - SEMESTER 1
curriculum{2,1} = {
    'WAR2101', 'Engineering Mathematics III', 4;
    'WAR2102', 'Engineering Geology', 4;
    'WAR2103', 'Hydrology I', 4;
    'WAR2104', 'Mechanics of Materials', 4;
    'WAR2105', 'Soil Mechanics', 4;
    'WAR2106', 'Computer Programming', 4
    };
% YEAR 2 - SEMESTER 2
curriculum{2,2} = {
    'WAR2201', 'Theory of Structure', 4;
    'WAR2202', 'Electrical Devices and Machines', 4;
    'WAR2203', 'Water Treatment', 4;
    'WAR2204', 'Engineering Hydrolics', 4;
    'WAR2205', 'Material Science', 4;
    'WAR2206', 'Hydrology II', 4
    };
% YEAR 3 - SEMESTER 1
curriculum{3,1} = {
    'WAR3101', 'Design of Water Supply Systems', 4;
    'WAR3102', 'GIS and Remote sensing', 4;
    'WAR3103', 'Waste Water Treatment', 4;
    'WAR3104', 'Hydrology III', 4;
    'WAR3105', 'Foundation Engineering', 4;
    'WAR3106', 'Quantitative Surveying', 4
    };
% YEAR 3 - SEMESTER 2
curriculum{3,2} = {
    'WAR3201', 'Hydro Information and Smart Water Systems', 4;
    'WAR3202', 'Ground Water Development And Management', 4;
    'WAR3203', 'Design of Steel and Concrete Structures', 4;
    'WAR3204', 'Irrigation and Drainage Engineering', 4;
    'WAR3205', 'Researcyh Methods And Engineering Project I', 3;
    'WAR3206', 'Entrepreneurship And Business Management', 3
    };
% YEAR 4 - SEMESTER 1
curriculum{4,1} = {
    'WAR4101', 'Engineering Project I', 2;
    'WAR4102', 'Planning and Development of Hydropower', 4;
    'WAR4103', 'Public Health and Sanitation Engineering', 4;
    'WAR4104', 'Engineering Ethics', 4;
    'WAR4105', 'Integrated Water Resources Management', 3;
    'WAR4106', 'Urban Drainage', 3
    };
% YEAR 4 - SEMESTER 2
curriculum{4,2} = {
    'WAR4201', 'Engineering Project II', 4;
    'WAR4202', 'Construction Management', 4;
    'WAR4203', 'Design of Hydropower Structures', 4;
    'WAR4204', 'Environmental Engineering', 3;
    'WAR4205', 'Intergrated Water Resource Management', 3;
    'WAR4206', 'Renewable Energy Technologies', 3
    };

% SELECT WHAT TO CALCULATE
fprintf('What would you like to calculate?\n\n');
fprintf('1. Current semester GPA only\n');
fprintf('2. GPA and CGPA up to the current semester\n');
fprintf('3. Exit\n\n');
choice = readNumber('Enter your choice (1-3): ', 1, 3);
if choice == 3
    fprintf('\nProgram terminated.\n');
    return;
end
    % store results
semesterGPA = NaN(4,2);
yearGPA = NaN(4,1);
totalWeightedPoints = 0;
totalCreditUnits = 0;
    % determine which semester to process
if choice == 1
    startYear = yearOfStudy;
    endYear = yearOfStudy;
    startSemester = currentSemester;
    endSemester = currentSemester;
else
    startYear = 1;
    endYear = yearOfStudy;
    startSemester = 1;
    endSemester = currentSemester;
end
% process each semeter
for y = startYear:endYear
    if y == startYear
        firstSemester = startSemester;
    else
        firstSemester = 1;
    end
    if y == endYear
        lastSemester = endSemester;
    else
        lastSemester = 2;
    end
    for s = firstSemester:lastSemester
        courses = curriculum{y,s};
        fprintf('\n');
        fprintf('YEAR %d - SEMESTER %d\n', y, s);
        fprintf('============================================================\n');
        fprintf('Enter marks for the following courses.\n');
        fprintf('Press ENTER without entering a mark to record ZERO.\n\n');
        numCourses = size(courses,1);
        marks = zeros(numCourses,1);
        gradePoints = zeros(numCourses,1);
        weightedPoints = zeros(numCourses,1);
        grades = cell(numCourses,1);
        semesterWeightedPoints = 0;
        semesterCreditUnits = 0;
        for c = 1:numCourses
            courseCode = courses{c,1};
            courseName = courses{c,2};
            creditUnit = courses{c,3};
            prompt = sprintf('%s - %s (%g CU): ', ...
                courseCode, courseName, creditUnit);
            marks(c) = readMark(prompt);
            [grades{c}, gradePoints(c)] = convertMark(marks(c));
            weightedPoints(c) = creditUnit * gradePoints(c);
            semesterWeightedPoints = ...
                semesterWeightedPoints + weightedPoints(c);
            semesterCreditUnits = ...
                semesterCreditUnits + creditUnit;
        end
% CALCULATE SEMESTER GPA
        semesterGPA(y,s) = semesterWeightedPoints / semesterCreditUnits;
        if choice == 2
            totalWeightedPoints = totalWeightedPoints + semesterWeightedPoints;
            totalCreditUnits = totalCreditUnits + semesterCreditUnits;
        end
        % display semester results
        fprintf('\n');
        fprintf('YEAR %d - SEMESTER %d RESULTS\n', y, s);
        fprintf('------------------------------------------------------------\n');
        fprintf('%-10s %-40s %6s %6s %8s %8s\n','Code', 'Course', 'CU', 'Mark', 'Grade', 'GP');
        fprintf('------------------------------------------------------------\n');
        for c = 1:numCourses
            fprintf('%-10s %-40s %6g %6.1f %8s %8.1f\n', ...
                courses{c,1}, ...
                courses{c,2}, ...
                courses{c,3}, ...
                marks(c), ...
                grades{c}, ...
                gradePoints(c));
        end
        fprintf('------------------------------------------------------------\n');
        fprintf('Total Credit Units: %.0f\n', semesterCreditUnits);
        fprintf('Total Weighted Grade Points: %.2f\n', semesterWeightedPoints);
        fprintf('SEMESTER GPA: %.2f\n', semesterGPA(y,s));
        fprintf('------------------------------------------------------------\n');
    end
% CALCULATE YEAR GPA
    yearWeightedPoints = 0;
    yearCreditUnits = 0;
    for s = 1:2
        if ~isnan(semesterGPA(y,s))
            courses = curriculum{y,s};
            numCourses = size(courses,1);
            semesterCU = 0;
            for c = 1:numCourses
                semesterCU = semesterCU + courses{c,3};
            end
            yearWeightedPoints = ...
                yearWeightedPoints + ...
                semesterGPA(y,s) * semesterCU;
            yearCreditUnits = ...
                yearCreditUnits + semesterCU;
        end
    end
    if yearCreditUnits > 0
        yearGPA(y) = yearWeightedPoints / yearCreditUnits;
    end
end

% CALCULATE AND DISPLAY FINAL RESULTS
fprintf('\n\n\n');
fprintf('                    FINAL RESULTS\n');
fprintf('============================================================\n');
fprintf('Student Name       : %s\n', studentName);
fprintf('Registration Number: %s\n', regNo);
fprintf('Programme          : %s\n', program);
fprintf('Current Year       : %d\n', yearOfStudy);
fprintf('Current Semester   : %d\n', currentSemester);
for y = 1:yearOfStudy
    for s = 1:2
        if ~isnan(semesterGPA(y,s))
            fprintf('Year %d Semester %d GPA: %.2f\n', ...
                y, s, semesterGPA(y,s));
        end
    end
end
fprintf('                       YEAR GPAs\n');
fprintf('------------------------------------------------------------\n');
for y = 1:yearOfStudy
    if ~isnan(yearGPA(y))
        fprintf('Year %d GPA: %.2f\n', y, yearGPA(y));
    end
end
if choice == 2
    CGPA = totalWeightedPoints / totalCreditUnits;
    fprintf('OVERALL CGPA: %.2f\n', CGPA);
    fprintf('CLASSIFICATION: %s\n', classifyCGPA(CGPA));
    fprintf('------------------------------------------------------------\n');
else
    fprintf('\nCGPA was not calculated because you selected GPA only.\n');
end

% GRADE
function value = readNumber(prompt, minimum, maximum)
    while true
        userInput = input(prompt, 's');
        value = str2double(userInput);
        if ~isnan(value) && value >= minimum && value <= maximum
            break;
        end
        fprintf('Invalid input. Enter a value from %g to %g.\n', ...
            minimum, maximum);
    end
end
function mark = readMark(prompt)
    while true
        userInput = input(prompt, 's');
        if isempty(strtrim(userInput))
            mark = 0;
            return;
        end
        mark = str2double(userInput);
        if ~isnan(mark) && mark >= 0 && mark <= 100
            return;
        end
        fprintf('Invalid mark. Enter a value between 0 and 100.\n');
        fprintf('Press ENTER without a mark to record ZERO.\n');
    end
end
function [grade, gradePoint] = convertMark(mark)
    if mark >= 80
        grade = 'A';
        gradePoint = 5.0;
    elseif mark >= 75
        grade = 'B+';
        gradePoint = 4.5;
    elseif mark >= 70
        grade = 'B';
        gradePoint = 4.0;
    elseif mark >= 65
        grade = 'B-';
        gradePoint = 3.5;
    elseif mark >= 60
        grade = 'C+';
        gradePoint = 3.0;
    elseif mark >= 55
        grade = 'D+';
        gradePoint = 2.5;
    elseif mark >= 50
        grade = 'D';
        gradePoint = 2.0;
    else
        grade = 'F';
        gradePoint = 0.0;
    end
end
function classification = classifyCGPA(CGPA)
    if CGPA >= 4.40
        classification = 'FIRST CLASS';
    elseif CGPA >= 3.60
        classification = 'SECOND CLASS UPPER';
    elseif CGPA >= 2.80
        classification = 'SECOND CLASS LOWER';
       
    elseif CGPA >= 2.00
        classification = 'PASS (THIRD CLASS)';
    else
        classification = 'FAIL';
    end
end
end

%% ASSIGNMENT 3: LEAF IMAGE PROCESSING
function leafProcessing()
clc;
% Leaf 1 - Lemon
leaf(1).plant = 'Lemon';
leaf(1).shape = 'Ovate';
leaf(1).margin = 'Serrated';
leaf(1).venation = 'Pinnate';
leaf(1).apex = 'Acute';
leaf(1).base = 'Cuneate';
leaf(1).arrangement = 'Alternate';
leaf(1).type = 'Simple';
leaf(1).texture = 'Smooth';

% Leaf 2 - Guava
leaf(2).plant = 'Guava';
leaf(2).shape = 'Elliptic';
leaf(2).margin = 'Entire';
leaf(2).venation = 'Pinnate';
leaf(2).apex = 'Acute';
leaf(2).base = 'Cuneate';
leaf(2).arrangement = 'Opposite';
leaf(2).type = 'Simple';
leaf(2).texture = 'Smooth';

% Leaf 3 - Mango
leaf(3).plant = 'Mango';
leaf(3).shape = 'Lanceolate';
leaf(3).margin = 'Entire';
leaf(3).venation = 'Pinnate';
leaf(3).apex = 'Acuminate';
leaf(3).base = 'Cuneate';
leaf(3).arrangement = 'Alternate';
leaf(3).type = 'Simple';
leaf(3).texture = 'Smooth';

% Leaf 4 - Maize
leaf(4).plant = 'Maize';
leaf(4).shape = 'Linear';
leaf(4).margin = 'Entire';
leaf(4).venation = 'Parallel';
leaf(4).apex = 'Acute';
leaf(4).base = 'Cuneate';
leaf(4).arrangement = 'Alternate';
leaf(4).type = 'Simple';
leaf(4).texture = 'Smooth';

% Leaf 5 - Cassava
leaf(5).plant = 'Cassava';
leaf(5).shape = 'Palmate-lobed';
leaf(5).margin = 'Lobed';
leaf(5).venation = 'Palmate';
leaf(5).apex = 'Acute';
leaf(5).base = 'Cordate';
leaf(5).arrangement = 'Alternate';
leaf(5).type = 'Simple';
leaf(5).texture = 'Smooth';

% IMAGE FILE NAMES
imageNames = {
    'leaf1.jpeg'
    'leaf2.jpg'
    'leaf3.jpg'
    'leaf4.jpg'
    'leaf5.jpg'
};

% READ AND PROCESS IMAGES
for i = 1:5
    leaf(i).originalImage = imread(imageNames{i});
    leaf(i).grayImage = rgb2gray(leaf(i).originalImage);
    leaf(i).binaryImage = imbinarize(leaf(i).grayImage);
    leaf(i).edgeImage = edge(leaf(i).grayImage);

    % MEASURE THE LEAF
    cleanImage = bwareaopen(leaf(i).binaryImage,500);
    measurements = regionprops(cleanImage, ...
        'Area','Perimeter','MajorAxisLength','MinorAxisLength');
    if ~isempty(measurements)
        [~,n] = max([measurements.Area]);
        leaf(i).area = measurements(n).Area;
        leaf(i).perimeter = measurements(n).Perimeter;
        leaf(i).length = measurements(n).MajorAxisLength;
        leaf(i).width = measurements(n).MinorAxisLength;
    else
        leaf(i).area = 0;
        leaf(i).perimeter = 0;
        leaf(i).length = 0;
        leaf(i).width = 0;
    end
end

% DISPLAY LEAF INFORMATION
for i = 1:5
    fprintf('\n\nLeaf %d: %s\n\n',i,leaf(i).plant);
    fprintf('Shape: %s\n',leaf(i).shape);
    fprintf('Margin: %s\n',leaf(i).margin);
    fprintf('Venation: %s\n',leaf(i).venation);
    fprintf('Apex: %s\n',leaf(i).apex);
    fprintf('Base: %s\n',leaf(i).base);
    fprintf('Arrangement: %s\n',leaf(i).arrangement);
    fprintf('Type: %s\n',leaf(i).type);
    fprintf('Texture: %s\n',leaf(i).texture);
   
    fprintf('\nMeasurements:\n\n');
    fprintf('Area: %.2f pixels\n',leaf(i).area);
    fprintf('Perimeter: %.2f pixels\n',leaf(i).perimeter);
    fprintf('Length: %.2f pixels\n',leaf(i).length);
    fprintf('Width: %.2f pixels\n',leaf(i).width);
end

% DISPLAY IMAGES
for i = 1:5
    figure;
    subplot(2,2,1);
    imshow(leaf(i).originalImage);
    title(['Original - ' leaf(i).plant]);

    subplot(2,2,2);
    imshow(leaf(i).grayImage);
    title('Grayscale');

    subplot(2,2,3);
    imshow(leaf(i).binaryImage);
    title('Binary');

    subplot(2,2,4);
    imshow(leaf(i).edgeImage);
    title('Edge');
end
end

%% ASSIGNMENT 4: NUMERICAL ANALYSIS
function numericalAnalysis()
clc;
% NEWTON-RAPHSON METHOD
fprintf('\n\n        NEWTON-RAPHSON METHOD\n');
fprintf('========================================\n');
f = @(x) x^3 + 7*x^2 + 5*x - 4;
df = @(x) 3*x^2 + 14*x + 5;
x = 0;
tol = 0.00001;
maxIter = 100;
for i = 1:maxIter
    xNew = x - f(x)/df(x);
    fprintf('Iteration %d: x = %.6f\n', i, xNew);
    if abs(xNew - x) < tol
        break;
    end
    x = xNew;
end
fprintf('\nRoot = %.6f\n', xNew);

% SECANT METHOD
fprintf('             SECANT METHOD\n');
fprintf('========================================\n');
f = @(x) x^3 + 7*x^2 + 5*x - 4;
x0 = 0;
x1 = 1;
tol = 0.00001;
maxIter = 100;
for i = 1:maxIter
    x2 = x1 - f(x1)*(x1-x0)/(f(x1)-f(x0));
    fprintf('Iteration %d: x = %.6f\n', i, x2);
    if abs(x2-x1) < tol
        break;
    end
    x0 = x1;
    x1 = x2;
end
fprintf('\nRoot = %.6f\n', x2);

% RUNGE-KUTTA 4TH ORDER METHOD
fprintf('       RUNGE-KUTTA 4TH ORDER\n');
fprintf('========================================\n');
% Differential equation: dy/dx = x^3 + 7*x^2 + 5*x - 4
% Initial condition: y(0) = 1
f = @(x,y) x^3 + 7*x^2 + 5*x - 4;
x = 0;
y = 1;
h = 0.1;
xEnd = 1;
fprintf('\n');
while x < xEnd
    k1 = h*f(x,y);
    k2 = h*f(x + h/2, y + k1/2);
    k3 = h*f(x + h/2, y + k2/2);
    k4 = h*f(x + h, y + k3);
    y = y + (k1 + 2*k2 + 2*k3 + k4)/6;
    x = x + h;
    fprintf('x = %.1f, y = %.6f\n', x, y);
end
fprintf('\nFinal answer:\n');
fprintf('y(%.1f) = %.6f\n', x, y);
end
