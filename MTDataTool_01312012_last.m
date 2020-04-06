function [] = MTDataTool()
% -------------------------------------------------------------------------

% MT Data Tool
% ------------------------------------
% GLOBAL VARIABLES

% -------------------------------------------------------------------------
% VARIABLE                              % DESCRIPTION
% -------------------------------------------------------------------------
UI = [];                                % user input structure
G  = [];                                % constant variables
MT = [];                                % maze transform data structure

mrgn = 10;                              % standard pixel margin

% -------------------------------------------------------------------------

% MT Data Tool
% ------------------------------------
% OBJECT HANDLES

% -------------------------------------------------------------------------
% ------------------------------------
% Figure Window
% ------------------------------------
% -------------------------------------------------------------------------
ScreenSize = get(0,'ScreenSize');
wH = 461; % window height
wW = 770; % window width
window = figure(...
    'units','pixels',...
    'position',[(ScreenSize(3)-wW)/2 (ScreenSize(4)-wH)/2 wW wH],...
    'menubar','none',...
    'name','Maze Transform Cell Data Tool',...
    'numbertitle','off',...
    'resize','off');
% -------------------------------------------------------------------------
% ------------------------------------
% "Open MT File" Menu Item
% ------------------------------------
% -------------------------------------------------------------------------
MenuBar.openLaps = uimenu(...
    'Parent',window,...
    'Label','Open MT File(s)',...
    'Callback',@openLaps_Call);
% -------------------------------------------------------------------------
    function [] = openLaps_Call(varargin)
        % open standard "open" dialog box
        [name,path] = uigetfile('*.xls','Select lap file(s)','Multiselect','on');
        % if user selects a file
        if ~isnumeric(name)
            % if user selects only one file, put "name" string into cell
            if ischar(name)
                name = {name};
            end
            
            % enter "name" and "path" into UI structure
            UI.name = name;
            UI.path = path;
            
            G.numFiles = length(name);
            G.lapStr = cell(G.numFiles,1);
            
            % get lap numbers from file names
            noLap = 100;
            for n = 1:G.numFiles
                ind = strfind(UI.name{n},'Lap');
                if isempty(ind)
                    ind = strfind(UI.name{n},'lap');
                end
                if ~isempty(ind)
                    G.lapNum(n,1) = str2double(UI.name{n}(ind+3:ind+4));
                    if isnan(G.lapNum(n,1))
                        G.lapNum(n,1) = str2double(UI.name{n}(ind+3));
                    end
                else
                    % if search can't find "Lap" in filename, number from 100
                    noLap = noLap+1;
                    G.lapNum(n,1) = noLap;
                end
            end
            
            % sort file names by lap number
            G.lapNum(:,2) = 1:G.numFiles;
            G.lapNum = sortrows(G.lapNum);
            UI.name = UI.name(G.lapNum(:,2));
            G.lapNum(:,2) = [];
            
            for n = 1:G.numFiles
                G.lapStr(n,1) = {num2str(G.lapNum(n))};
            end
            
            % import excel data
            [MT,G] = importDataToolData(UI,G,MT);
            
            % set constants
            G.numCells = length(G.cellName);
            G.binW = G.bins(2);
            G.xTickInc = floor(length(G.bins)/7)*G.binW;
            G.start = 1;
            G.end = length(G.bins);
            A.xTick = [0,G.xTickInc:G.xTickInc:G.bins(end)-G.xTickInc,G.bins(end)];
            A.freqMax = 0;
            
            % set cell selection (on cell list)
            G.cellSel = 1:G.numCells;
            
            % set cell list to cell names
            set(C.list,'String',G.cellName);
        end
    end
% -------------------------------------------------------------------------
% ------------------------------------
% Select Cells for Analysis
% ------------------------------------
% -------------------------------------------------------------------------
C.menu.parent = uicontextmenu(...
    'Parent',window);
% ------------------------------------
% Reset cell list
C.menu.reset = uimenu(...
    'Parent',C.menu.parent,...
    'Label','Reset',...
    'Callback',@cellsResetCall);
% ------------------------------------
    function [] = cellsResetCall(varargin)
        if ~isempty(G)
            set(C.list,'String',G.cellName);
            set(C.list,'Val',1);
            G.cellSel = 1:G.numCells;
        end
    end
% ------------------------------------
% Narrow cell list to selected cells
C.menu.confirm = uimenu(...
    'Parent',C.menu.parent,...
    'Label','Confirm',...
    'Callback',@cellsConfirmCall);
% ------------------------------------
    function [] = cellsConfirmCall(varargin)
        if ~isempty(G)
            choice = get(C.list,{'String','Val'});
            if ~isempty(choice{1})
                choice{1} = choice{1}(choice{2});
                set(C.list,'String',choice{1},'Val',1:length(choice{1}));
                % re-calculating UI.CellSel
                % G.cellSel = G.cellName index of user selected cells
                numCellsSel = length(choice{1});
                match = zeros(numCellsSel,1);
                for n = 1:numCellsSel
                    match(n) = find(strcmp(choice{1}{n},G.cellName),1);
                end
                G.cellSel = match;
            end
        end
    end
% ------------------------------------
% Delete selected cells from list
C.menu.delete = uimenu(...
    'Parent',C.menu.parent,...
    'Label','Delete',...
    'Callback',@cellsDeleteCall);
% ------------------------------------
    function [] = cellsDeleteCall(varargin)
        if ~isempty(G)
            choice = get(C.list,{'String','Val'});
            if ~isempty(choice{1})
                choice{1}(choice{2}) = [];
                if choice{2}(1) ~= length(choice{1})+1;
                    newVal = choice{2}(1);
                else
                    newVal = length(choice{1});
                end
                set(C.list,'String',choice{1},'Val',newVal);
                % re-calculating UI.CellSel
                % G.cellSel = G.cellName index of user selected cells
                numCellsSel = length(choice{1});
                match = zeros(numCellsSel,1);
                for n = 1:numCellsSel
                    match(n) = find(strcmp(choice{1}{n},G.cellName),1);
                end
                G.cellSel = match;
            end
        end
    end
% ------------------------------------
C.pW = 280; % panel width
C.pH = 185; % panel height
C.panel = uipanel(...
    'Parent',window,...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',12,...
    'FontWeight','bold',...
    'TitlePosition','lefttop',...
    'Title','Select Cells For Analysis',...
    'BackgroundColor',[.8 .8 .8],...
    'BorderType','beveledout',...
    'Position',[mrgn, wH-mrgn-C.pH, C.pW, C.pH]);
% ------------------------------------
C.lW = C.pW-2*mrgn;
C.lH = C.pH-2*mrgn-mrgn-17; % 17 for text field
C.list = uicontrol(...
    'Parent',C.panel,...
    'Style','list',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'String',{'Cell Names'},...
    'Min',0,'Max',100,...
    'BackgroundColor','white',...
    'UIContextMenu',C.menu.parent,...
    'Position',[mrgn, C.pH-8-mrgn-C.lH, C.lW, C.lH],...
    'Callback',@plotCellCall);
% ------------------------------------
pCCc = 0;
    function [] = plotCellCall(varargin)
        if ~isempty(G)
            SelType = get(window,'SelectionType');
            if strcmp(SelType,'open')
                pCCc = pCCc+1;
                if pCCc == 1
                    setLineLabelsVisible;
                end
                G.currCell = G.cellSel(get(C.list,'Val'));
                A.freq = mean(MT{G.currCell}.freq,2);
                A.freqMax = max(A.freq);
                set(A.label,'String',MT{G.currCell}.name);
                fixAxes;
            end
        end
    end
% ------------------------------------
    function [] = setLineLabelsVisible(varargin)
        set(A.start.bin,'Visible','on','String',num2str(G.bins(G.start)));
        set(A.start.label,'Visible','on');
        set(A.start.color,'Visible','on');
        
        set(A.end.bin,'Visible','on','String',num2str(G.bins(G.end)));
        set(A.end.label,'Visible','on');
        set(A.end.color','Visible','on');
    end
% ------------------------------------
    function [] = fixAxes(varargin)
        if ~isempty(G);
            freq = sprintf('%2.2g',A.freqMax);
            freq = str2double(freq);
            A.yTick = unique(sortrows([1;freq]))';
            A.yLim = [0,1.2*A.freqMax];
            A.xLim = [G.bins(1)-G.binW/2,G.bins(end)+G.binW/2];
            set(A.axes,...
                'XTick',A.xTick,...
                'YTick',A.yTick,...
                'TickDir','out',...
                'TickLength',[0.005,0.005],...
                'Box','off',...
                'XLim',A.xLim,...
                'YLim',A.yLim);
            set(A.freqBar,...
                'Visible','on',...
                'XData',G.bins,...
                'YData',A.freq);
            set(A.start.line,...
                'Visible','on',...
                'XData',[G.bins(G.start),G.bins(G.start)],...
                'YData',A.yLim);
            set(A.end.line,...
                'Visible','on',...
                'XData',[G.bins(G.end),G.bins(G.end)],...
                'YData',A.yLim);
        else
            set(A.freqBar,'Visible','off');
            set(A.start.line,'Visible','off');
            set(A.end.line,'Visible','off');
            set(A.axes,'XTick',[],'YTick',[]);
        end
    end
% ------------------------------------
C.note = uicontrol(...
    'Parent',C.panel,...
    'Style','text',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'HorizontalAlignment','left',...
    'String','Double click cell name to plot',...
    'BackgroundColor',[.8,.8,.8],...
    'Position',[mrgn, C.pH-8-mrgn-C.lH-22, C.pW-2*mrgn, 17]);
% -------------------------------------------------------------------------
% ------------------------------------
% Cell Data Analysis
% ------------------------------------
% -------------------------------------------------------------------------
DA.panel = uipanel(...
    'Parent',window,...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',12,...
    'FontWeight','bold',...
    'TitlePosition','lefttop',...
    'Title','Cell Data Analysis',...
    'BackgroundColor',[.8 .8 .8],...
    'BorderType','beveledout');
% ------------------------------------
DA.printCell.check = uicontrol(...
    'Parent',DA.panel,...
    'Style','checkbox',...
    'BackgroundColor',[.8 .8 .8],...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'String','Cell Sheets');
% ------------------------------------
DA.printAgg.check = uicontrol(...
    'Parent',DA.panel,...
    'Style','checkbox',...
    'BackgroundColor',[.8 .8 .8],...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'String','Aggregate Sheet');
% ------------------------------------
DA.applyFilters.check = uicontrol(...
    'Parent',DA.panel,...
    'Style','checkbox',...
    'BackgroundColor',[.8 .8 .8],...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'String','Apply filters to cell anaylysis');
% ------------------------------------
DA.button = uicontrol(...
    'Parent',DA.panel,...
    'Style','pushbutton',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'String','Run Analysis',...
    'BackgroundColor',[.8 .8 .8],...
    'Callback',@runCellDataAnalysis);
% ------------------------------------
    function [] = runCellDataAnalysis(varargin)
        if ~isempty(G)
            % in case start/end are implemented
            UI.printAggregateSheet = get(DA.printAgg.check,'Val');
            UI.printCellSheet = get(DA.printCell.check,'Val');
            UI.freqFilter = get(F.freq.check,'Val');
            UI.timeFilter = get(F.time.check,'Val');
            UI.applyFilters = get(DA.applyFilters.check,'Val');
            if UI.applyFilters
                filteredCells = filterCells(UI,MT,G);
            else
                filteredCells = MT;
            end
            CellDataAnalysis(UI,filteredCells,G);
        end
    end

% Set DA.panel Positions
% ------------------------------------
mrgn = 10;
DA.buttonHeight = 20;
DA.checkHeight = 17;
DA.pW = C.pW;
DA.pH = mrgn;
set(DA.button,'position',[mrgn,DA.pH,DA.pW-2*mrgn,DA.buttonHeight]);
DA.pH = DA.pH+DA.buttonHeight+mrgn/2;
set(DA.applyFilters.check,'position',[mrgn,DA.pH,DA.pW-2*mrgn,DA.checkHeight]);
DA.pH = DA.pH+DA.checkHeight+mrgn/2;
set(DA.printAgg.check,'position',[mrgn,DA.pH,DA.pW-2*mrgn,DA.checkHeight]);
DA.pH = DA.pH+DA.checkHeight+mrgn/2;
set(DA.printCell.check,'position',[mrgn,DA.pH,DA.pW-2*mrgn,DA.checkHeight]);
DA.pH = DA.pH+DA.checkHeight+mrgn+8;
set(DA.panel,'position',[mrgn,wH-C.pH-DA.pH-2*mrgn,DA.pW,DA.pH]);
% -------------------------------------------------------------------------
% ------------------------------------
% Axes
% ------------------------------------
% -------------------------------------------------------------------------
A.pW = wW-C.pW-3*mrgn;
A.pH = C.pH+DA.pH+mrgn-8;
A.panel = uipanel(...
    'Parent',window,...
    'Units','pixels',...
    'BackgroundColor',[.8 .8 .8],...
    'BorderType','beveledout',...
    'Position',[2*mrgn+C.pW, wH-mrgn-8-A.pH, A.pW, A.pH]);
% ------------------------------------
A.aW = A.pW-4*mrgn;
A.aH = A.aW/2;
A.axes = axes(...
    'Parent',A.panel,...
    'Units','pixels',...
    'XTick',[],'YTick',[],...
    'FontUnits','pixels',...
    'FontSize',10,...
    'Position',[3*mrgn, A.pH-1*mrgn-A.aH, A.aW, A.aH]);
% ------------------------------------
hold on;
% ------------------------------------
A.freqBar = bar(...
    0,0,...
    'BarWidth',1,...
    'Parent',A.axes);
% ------------------------------------
A.start.line = line(...
    'XData',0,...
    'YData',0,...
    'Parent',A.axes,...
    'LineWidth',2,...
    'Color',[0 1 0],...
    'Visible','off',...
    'ButtonDownFcn',@startDragStart);
% ------------------------------------
    function [] = startDragStart(varargin)
        set(window,'WindowButtonMotionFcn',@dragStart);
        set(window,'WindowButtonUpFcn',@stopDragStart);
    end
% ------------------------------------
    function [] = dragStart(varargin)
        cursor = get(A.axes,'CurrentPoint');
        cursor = round(cursor(1)/G.binW);
        if cursor < 1, cursor = 1; end
        if cursor > length(G.bins), cursor = length(G.bins); end
        G.start = cursor;
        set(A.start.line,'XData',[G.bins(cursor),G.bins(cursor)]);
        set(A.start.bin,'String',num2str(G.bins(cursor)));
    end
% ------------------------------------
    function [] = stopDragStart(varargin)
        set(window,'WindowButtonMotionFcn','');
    end
% ------------------------------------
A.start.bin = uicontrol(...
    'Parent',A.panel,...
    'Visible','off',...
    'Style','text',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontWeight','normal',...
    'FontSize',12,...
    'HorizontalAlignment','right',...
    'String','',...
    'BackgroundColor',[.8,.8,.8],...
    'Position',[A.pW-mrgn-28-mrgn-12-2-2*28, mrgn, 28, 17]);
% ------------------------------------
A.start.label = uicontrol(...
    'Parent',A.panel,...
    'Visible','off',...
    'Style','text',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontWeight','bold',...
    'FontSize',12,...
    'HorizontalAlignment','left',...
    'String','Start',...
    'BackgroundColor',[.8,.8,.8],...
    'Position',[A.pW-mrgn-2*28-mrgn-12-2-2*28, mrgn, 28, 17]);
% ------------------------------------
A.start.color = uipanel(...
    'Parent',A.panel,...
    'Visible','off',...
    'Units','pixels',...
    'BackgroundColor',[0 1 0],...
    'BorderType','beveledout',...
    'Position',[A.pW-mrgn-12-2*28-2-mrgn-12-2-2*28, mrgn+17/2-1, 12, 4]);
% ------------------------------------
A.end.line = line(...
    'XData',0,...
    'YData',0,...
    'Parent',A.axes,...
    'LineWidth',2,...
    'Color',[1 0 0],...
    'Visible','off',...
    'ButtonDownFcn',@startDragEnd);
% ------------------------------------
    function [] = startDragEnd(varargin)
        set(window,'WindowButtonMotionFcn',@dragEnd);
        set(window,'WindowButtonUpFcn',@stopDragEnd);
    end
% ------------------------------------
    function [] = dragEnd(varargin)
        cursor = get(A.axes,'CurrentPoint');
        cursor = round(cursor(1)/G.binW);
        if cursor < 1, cursor = 1; end
        if cursor > length(G.bins), cursor = length(G.bins); end
        G.end = cursor;
        set(A.end.line,'XData',[G.bins(cursor),G.bins(cursor)]);
        set(A.end.bin,'String',num2str(G.bins(cursor)));
    end
% ------------------------------------
    function [] = stopDragEnd(varargin)
        set(window,'WindowButtonMotionFcn','');
    end
% ------------------------------------
A.end.bin = uicontrol(...
    'Parent',A.panel,...
    'Visible','off',...
    'Style','text',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontWeight','normal',...
    'FontSize',12,...
    'HorizontalAlignment','right',...
    'String','',...
    'BackgroundColor',[.8,.8,.8],...
    'Position',[A.pW-mrgn-28, mrgn, 28, 17]);
% ------------------------------------
A.end.label = uicontrol(...
    'Parent',A.panel,...
    'Visible','off',...
    'Style','text',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontWeight','bold',...
    'FontSize',12,...
    'HorizontalAlignment','left',...
    'String','End',...
    'BackgroundColor',[.8,.8,.8],...
    'Position',[A.pW-mrgn-2*28, mrgn, 28, 17]);
% ------------------------------------
A.end.color = uipanel(...
    'Parent',A.panel,...
    'Visible','off',...
    'Units','pixels',...
    'BackgroundColor',[1 0 0],...
    'BorderType','beveledout',...
    'Position',[A.pW-mrgn-12-2*28-2, mrgn+17/2-1, 12, 4]);
% ------------------------------------
A.end.line = line(...
    'XData',0,...
    'YData',0,...
    'Parent',A.axes,...
    'LineWidth',2,...
    'Color',[1 0 0],...
    'Visible','off',...
    'ButtonDownFcn',@startDragEnd);
% ------------------------------------
hold off;
% ------------------------------------
fixAxes;
% ------------------------------------
A.lineSep = uipanel(...
    'Parent',A.panel,...
    'Units','pixels',...
    'BackgroundColor',[.8 .8 .8],...
    'BorderType','beveledout',...
    'Position',[mrgn, mrgn+17+5, (A.pW-mrgn*2), 2]);
A.label = uicontrol(...
    'Parent',A.panel,...
    'Style','text',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontWeight','bold',...
    'FontSize',12,...
    'HorizontalAlignment','left',...
    'String','Empty Axes',...
    'BackgroundColor',[.8,.8,.8],...
    'Position',[mrgn, mrgn, (A.pW-mrgn*2)/2, 17]);
% -------------------------------------------------------------------------
% ------------------------------------
% Filters
% ------------------------------------
% -------------------------------------------------------------------------
F.panel = uipanel(...
    'Parent',window,...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',12,...
    'FontWeight','bold',...
    'TitlePosition','lefttop',...
    'Title','Filter Options',...
    'BackgroundColor',[.8 .8 .8],...
    'BorderType','beveledout');
% ------------------------------------
F.freq.check = uicontrol(...
    'Parent',F.panel,...
    'Style','checkbox',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'BackgroundColor',[.8,.8,.8],...
    'String','Frequency Filter (% max frequency)');
% ------------------------------------
F.freq.edit = uicontrol(...
    'Parent',F.panel,...
    'Style','edit',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'BackgroundColor',[1,1,1],...
    'String','20',...
    'Callback',@freqEditCall);
G.freqThresh = 0.20;
% ------------------------------------
    function [] = freqEditCall(varargin)
        num = str2double(get(F.freq.edit,'String'));
        if isnan(num), num = 20; end
        if num < 0, num = 0; end
        if num > 100, num = 100; end
        set(F.freq.edit,'String',num2str(num));
        G.freqThresh = num/100;       
    end
% ------------------------------------
F.freq.slide = uicontrol(...
    'Parent',F.panel,...
    'Style','slider',...
    'Units','pixels',...
    'BackgroundColor',[1 1 1],...
    'Callback',@freqSlideCall);
% ------------------------------------
    function [] = freqSlideCall(varargin)
    end
% ------------------------------------
F.time.check = uicontrol(...
    'Parent',F.panel,...
    'Style','checkbox',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'BackgroundColor',[.8,.8,.8],...
    'String','Time Filter (s)');
% ------------------------------------
F.time.edit = uicontrol(...
    'Parent',F.panel,...
    'Style','edit',...
    'Units','pixels',...
    'FontUnits','pixels',...
    'FontSize',10,...
    'BackgroundColor',[1,1,1],...
    'String','1',...
    'Callback',@timeEditCall);
G.timeThresh = 1;
% ------------------------------------
    function [] = timeEditCall(varargin)
        num = str2double(get(F.time.edit,'String'));
        if isnan(num), num = 1; end
        if num < 0, num = 0; end
        if num > 100, num = 100; end
        set(F.time.edit,'String',num2str(num));
        G.timeThresh = num;       
    end
% ------------------------------------
F.time.slide = uicontrol(...
    'Parent',F.panel,...
    'Style','slider',...
    'Units','pixels',...
    'BackgroundColor',[1 1 1],...
    'Callback',@timeSlideCall);
% ------------------------------------
    function [] = timeSlideCall(varargin)
    end
% ------------------------------------
F.printFilteredLapFiles = uicontrol(...
    'Parent',F.panel,...
    'Style','pushbutton',...
    'Units','pixels',...
    'BackgroundColor',[.8,.8,.8],...
    'FontUnits','pixels',...
    'FontWeight','normal',...
    'FontSize',10,...
    'String','Write Filtered Lap Files',...
    'Callback',@printFilteredLapFiles_Call);
% ------------------------------------
    function [] = printFilteredLapFiles_Call(varargin)
        if ~isempty(G)
            UI.freqFilter = get(F.freq.check,'Val');
            UI.timeFilter = get(F.time.check,'Val');
            filteredCells = filterCells(UI,MT,G);
            filteredLaps = cellsToLaps(filteredCells,G);
            printFilteredLaps(filteredLaps,UI.name,UI.path);
        end
    end
% ------------------------------------
mrgn = 10;
F.pW = C.pW;
F.checkHeight = 17;
F.buttonHeight = 20;
F.slideWidth = 11;
F.editWidth = 27;

F.pH = mrgn;
set(F.printFilteredLapFiles,'position',[mrgn,F.pH,F.pW-2*mrgn,F.buttonHeight]);

F.pH = F.pH+F.buttonHeight+mrgn/2;
set(F.time.slide,'position',[F.pW-mrgn-F.slideWidth,F.pH,F.slideWidth,F.checkHeight]);
set(F.time.edit,'position',[F.pW-mrgn-F.slideWidth-F.editWidth,F.pH,F.editWidth,F.checkHeight]);
set(F.time.check,'position',[mrgn,F.pH,F.pW-F.slideWidth-F.editWidth-2*mrgn,F.checkHeight]);

F.pH = F.pH+F.checkHeight+mrgn/2;
set(F.freq.slide,'position',[F.pW-mrgn-F.slideWidth,F.pH,F.slideWidth,F.checkHeight]);
set(F.freq.edit,'position',[F.pW-mrgn-F.slideWidth-F.editWidth,F.pH,F.editWidth,F.checkHeight]);
set(F.freq.check,'position',[mrgn,F.pH,F.pW-F.slideWidth-F.editWidth-2*mrgn,F.checkHeight]);

F.pH = F.pH+F.checkHeight+mrgn+8;
set(F.panel,'position',[mrgn,wH-C.pH-DA.pH-F.pH-3*10,F.pW,F.pH]);
% ------------------------------------
end % End of GUI
% -------------------------------------------------------------------------


% Filter Cell Data
% -------------------------------------------------------------------------
function [filteredCells] = filterCells(UI,unfilteredCells,G)
Cell = unfilteredCells(G.cellSel);
G.numCells = length(Cell);
G.numBins = length(G.bins);

% Zero Out Anything Outside of Start/End
if G.start < G.end
    if G.start ~= 1 || G.end ~= G.numBins
        for m = 1:G.numCells
            Cell{m}.freq(1:G.start-1,:) = zeros;
            Cell{m}.freq(G.end+1:end,:) = zeros;
        end
    end
else
    for m = 1:G.numCells
        Cell{m}.freq(G.end+1:G.start-1) = zeros;
    end
end

% Apply Time Filter
if UI.timeFilter
    [row,col] = find(G.time >= G.timeThresh);
    ind = sub2ind(size(G.time),row,col);
    for m = 1:G.numCells
        Cell{m}.freq(ind) = 0;
    end
end

% Apply Freq Filter
if UI.freqFilter
    for m = 1:G.numCells
        maxFreq = max(Cell{m}.freq);
        Cell{m}.freq(bsxfun(@lt,Cell{m}.freq,G.freqThresh*maxFreq)) = 0;
    end
end

filteredCells = unfilteredCells;
filteredCells(G.cellSel) = Cell;

end
% -------------------------------------------------------------------------


% Cells To Laps
% -------------------------------------------------------------------------
function [file] = cellsToLaps(Cell,G)
for m = 1:G.numFiles
    for n = 1:G.numCells
        file{m}.sheet{n}.data(:,1) = G.bins;
        file{m}.sheet{n}.data(:,2) = G.time(:,m);
        file{m}.sheet{n}.data(:,3) = Cell{n}.SC(:,m);
        file{m}.sheet{n}.data(:,4) = Cell{n}.freq(:,m);
        file{m}.sheet{n}.name = Cell{n}.name;
        file{m}.sheet{n}.data = num2cell(file{m}.sheet{n}.data);
    end
end
end
% -------------------------------------------------------------------------


% Print Filtered Laps
% -------------------------------------------------------------------------
function [] = printFilteredLaps(file,name,path)

excelObj = actxserver('Excel.Application');

numFiles = length(file);
numSheets = length(file{1}.sheet);

for m = 1:numFiles
    Workbook = excelObj.Workbooks.Add;
    
    for n = numSheets:-1:1
        % determine data range
        Size = size(file{m}.sheet{n}.data);
        OverShoot = Size(2)-26;
        if OverShoot > 0
            numAlphabets = ceil(Size(2)/26)-1;
            excelRange = ['a1:' char('a'+numAlphabets-1) char('a'+OverShoot-26*(numAlphabets-1)-1) num2str(Size(1))];
        else
            excelRange = ['a1:' char('a'+Size(2)-1) num2str(Size(1))];
        end
        % add sheet
        Worksheet{n} = excelObj.sheets.Add;
        % name sheet
        Worksheet{n}.Name = file{m}.sheet{n}.name;
        % get range
        Range{n} = get(Worksheet{n},'Range',excelRange);
        % print data
        Range{n}.Value = file{m}.sheet{n}.data;
    end
    
    % delete standard worksheets
    excelObj = deleteStdWorksheets(excelObj);
    
    % save workbook
    Workbook.SaveAs([path name{m}(1:end-4) '_filtered.xls']); Workbook.Close;
end

excelObj.Quit; delete(excelObj);

end
% -------------------------------------------------------------------------


function [excelObj] = deleteStdWorksheets(excelObj)
Worksheets = excelObj.sheets;
numSheets = Worksheets.Count;
sheetIdx = 1;
sheetIdx2 = 1;
while sheetIdx2 <= numSheets
   sheetName = Worksheets.Item(sheetIdx).Name(1:end-1);
   if strcmp(sheetName,'Sheet')
      Worksheets.Item(sheetIdx).Delete;
   else
      sheetIdx = sheetIdx + 1;
   end
   sheetIdx2 = sheetIdx2 + 1;
end
end


% Cell Data Analysis
% -------------------------------------------------------------------------
function [] = CellDataAnalysis(UI,Cell,G)
G.finalBin = G.bins(end);

G.numCells = length(G.cellSel);
Cell = Cell(G.cellSel);

if G.start < G.end
    G.numBins = G.end-G.start+1;
    G.bins = G.bins(G.start:G.end,:);
else
    G.numBins = (length(G.bins)-G.start+1)+G.end;
    G.bins = (G.bins(G.start):G.binW:(G.bins(G.start)+G.binW*G.numBins-G.binW))';
end

% align time data according the UI selected start/end bins
temp = selectData(G.time,G.start,G.end);
G.time = temp{1};
sheet = [];

for m = 1:G.numCells
    % align cell freq, SC data according the UI selected start/end bins
    temp = selectData({Cell{m}.SC,Cell{m}.freq},G.start,G.end);
    Cell{m}.SC = temp{1};
    Cell{m}.freq = temp{2};
    % center of mass
    Cell{m} = centerOfMass(Cell{m},G);
    % average frequency
    Cell{m} = averageFrequency(Cell{m},G);
    % # bins w/ freq > 0 hz
    Cell{m} = numFgtZero(Cell{m},G);
    % correlation
    Cell{m} = correlation(Cell{m},G);
    % info per spike
    Cell{m} = infoPerSpike(Cell{m},G);
    % other stats
    Cell{m} = otherStats(Cell{m},G);
    % aggregate stats
    if UI.printAggregateSheet
        O.CofM(:,m)         = Cell{m}.CofM;
        O.avgF(:,m)         = Cell{m}.avgF;
        O.infoPerSpike(:,m) = Cell{m}.infoPerSpike;
        O.shiftRPL(:,m)     = Cell{m}.shiftRPL;
        O.shiftAbs(:,m)     = Cell{m}.shiftAbs;
        O.shiftROP(:,m)     = Cell{m}.shiftROP;
        O.rRPL(:,m)         = Cell{m}.rRPL;
        O.rAvgAll(:,m)      = Cell{m}.rAvgAll;
        O.rAvg5L(:,m)       = Cell{m}.rAvg5L;
        O.numFgtZero(:,m)   = Cell{m}.numFgtZero;
    end
    
    if UI.printCellSheet
        sheet{m}.data = printCellData(Cell{m},G);
        sheet{m}.name = Cell{m}.name;
    end
end

if UI.printAggregateSheet
    aggregateSheet.data = printAggregateData(O,G);
    aggregateSheet.name = 'Aggregate Data';
    sheet = [sheet,aggregateSheet];
    if ~iscell(sheet), sheet = {sheet}; end
end

if UI.printAggregateSheet || UI.printCellSheet
    writeToExcel(sheet);
end

end
% -------------------------------------------------------------------------


% Select Data (used to realign with start/end bins)
% -------------------------------------------------------------------------
function [data] = selectData(data,s,e)
if ~iscell(data),data = {data}; end
if s < e
    for m = 1:length(data)
        data{m} = data{m}(s:e,:);
    end
else
    for m = 1:length(data)
        data{m} = [data{m}(s:end,:);data{m}(1:e,:)];
    end
end
end
% -------------------------------------------------------------------------


% Center of Mass
% -------------------------------------------------------------------------
function [Cell] = centerOfMass(Cell,G)
weight = bsxfun(@times,Cell.freq,G.bins);
Cell.CofM = (sum(weight)./sum(Cell.freq))';
end
% -------------------------------------------------------------------------


% Average Frequency
% -------------------------------------------------------------------------
function [Cell] = averageFrequency(Cell,G)
Cell.avgF = mean(Cell.freq)';
end
% -------------------------------------------------------------------------


% Correlation
% -------------------------------------------------------------------------
function [Cell] = correlation(Cell,G)
for i = 1:G.numFiles
    x = Cell.freq(:,i);
    for j = 1:G.numFiles
        y = Cell.freq(:,j);
        yShifted = zeros(G.numBins,G.numBins);
        for h = 1:G.numBins
            % shift forward by 1 bin
            yShifted(:,h) = [y(end-h+1:end);y(1:end-h)];
        end
        r = corrcoef([x,y,yShifted]);
        Cell.r(j,i) = r(2,1);
        % find max correlation
        rIndMax = find(r == max(r(2:end,1)),1);
        if ~isempty(rIndMax)
            % ShiftB = bins shifted at max corr
            Cell.shiftB(j,i) = rIndMax-1;
            % ShiftR = max corr
            Cell.shiftR(j,i) = r(rIndMax);
        else
            Cell.shiftB(j,i) = 0;
            Cell.shiftR(j,i) = 0;
        end
    end
end
Cell.r(isnan(Cell.r)) = 0;
Cell.shiftB(Cell.shiftB > G.numBins/2) = Cell.shiftB(Cell.shiftB > G.numBins/2)-G.numBins;
end
% -------------------------------------------------------------------------


% Information Per Spike (Spatial Information)
% -------------------------------------------------------------------------
function [Cell] = infoPerSpike(Cell,G)
TimeData = G.time;
Ri = Cell.freq;

R = mean(Ri);
sumTime = sum(TimeData);

Pi = bsxfun(@rdivide,TimeData,sumTime);
RiOverR = bsxfun(@rdivide,Ri,R);

temp = bsxfun(@times,bsxfun(@times,Pi,RiOverR),log2(RiOverR));
temp(isnan(temp)) = 0;

Cell.infoPerSpike = sum(temp)';
end
% -------------------------------------------------------------------------


% Number of Frequencies Greater Than Zero
% -------------------------------------------------------------------------
function [Cell] = numFgtZero(Cell,G)
[~,col] = find(Cell.freq > 0);
for m = 1:G.numFiles
    temp(m) = length(col(col == m));
end
Cell.numFgtZero = temp';
end
% -------------------------------------------------------------------------


% Other Stats
% -------------------------------------------------------------------------
function [Cell] = otherStats(Cell,G)
[r,c] = find(eye(size(Cell.shiftB))~=0);
row = r + 1; col = c; row(end) = []; col(end) = [];
Index = sub2ind(size(Cell.shiftB),row,col);

% shift relative to prior lap
Cell.shiftRPL = Cell.shiftB(Index);
% absolute shift
Cell.shiftAbs = abs(Cell.shiftRPL);
% shift relative to original position
for m = 1:length(Cell.shiftB(:,1))-1
    Cell.shiftROP(m) = sum(Cell.shiftRPL(1:m));
end
% correlation relative to prior lap
Cell.rRPL = Cell.r(Index);
Diag = sub2ind([G.numFiles,G.numFiles],r,c);
temp = Cell.r; temp(Diag) = 0;
% average correlation over all laps
Cell.rAvgAll = (sum(temp)/(G.numFiles-1))';
% average correlation over sets of five laps
ind = 0:5:G.numFiles;
if ind(end) ~= G.numFiles, ind(end+1) = G.numFiles; end
for m = 1:length(ind)-1
    num = ind(m+1)-(ind(m)+1);
    Cell.rAvg5L{m,1} = (sum(temp(ind(m)+1:ind(m+1),ind(m)+1:ind(m+1)))/num)';
end
Cell.rAvg5L = cell2mat(Cell.rAvg5L);
end
% -------------------------------------------------------------------------


% Print Cell Data
% -------------------------------------------------------------------------
function [Sheet] = printCellData(Cell,G)

SheetWidth = G.numFiles*2+1;
Sheet = cell(1,SheetWidth);
Sheet(1,1) = {Cell.name};
Sheet(end+1,:) = {[]};

Sheet(end+1,1:5) = {...
    'Lap Number','Center of Mass (cm)',...
    'Average Frequency (Hz)',...
    'Number of Bins w/ Frequency > 0 Hz',...
    'Information Per Spike'};
temp = [G.lapNum,Cell.CofM,Cell.avgF,Cell.numFgtZero,Cell.infoPerSpike];
[M,N] = size(temp);
Sheet(end+1:length(Sheet(:,1))+M,1:N) = num2cell(temp);
Sheet(end+1,:) = {[]};

Sheet(end+1,1) = {'Correlation'};
Sheet(end+1,2:G.numFiles+1) = num2cell(G.lapNum');
temp = [G.lapNum,Cell.r];
[M,N] = size(temp);
Sheet(end+1:length(Sheet(:,1))+M,1:N) = num2cell(temp);
Sheet(end+1,:) = {[]};

Sheet(end+1,1) = {'Frequency Data'};
Sheet(end+1,2:G.numFiles+1) = num2cell(G.lapNum');
temp = [G.bins,Cell.freq];
[M,N] = size(temp);
Sheet(end+1:length(Sheet(:,1))+M,1:N) = num2cell(temp);
Sheet(end+1,:) = {[]};

Sheet(end+1,1) = {'Shifting Correlation (given in number of Bins shifted)'};
Sheet(end+1,2:2:G.numFiles*2+1) = num2cell(G.lapNum');
Sheet(end+1,2:2:G.numFiles*2+1) = {'Shift'};
Sheet(end,3:2:G.numFiles*2+1) = {'r'};
a = 1:2:G.numFiles*2;
b = 2:2:G.numFiles*2;
clear temp;
temp(:,a) = Cell.shiftB;
temp(:,b) = Cell.shiftR;
temp = [G.lapNum,temp];
[M,N] = size(temp);
Sheet(end+1:length(Sheet(:,1))+M,1:N) = num2cell(temp);
clear temp; 

end
% -------------------------------------------------------------------------


% Print Aggregate Data
% -------------------------------------------------------------------------
function [sheet]= printAggregateData(O,G)
% constants
sheet = cell(1,G.numCells+4);
statsHeader = {'Mean';'StDev';'Count'};
RowHeader = cell(G.numFiles-1,1);
for n = 1:G.numFiles-1
    RowHeader{n} = ['L' num2str(G.lapNum(n)) '-L' num2str(G.lapNum(n+1))];
end

sheet(1,1) = {'AGGREAGATE DATA'};
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Center of Mass (cm)'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.CofM);
temp = getMeanStdCountRow(temp,true);
tempHeader = [num2cell(G.lapNum);statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Average Frequency (Hz)'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.avgF);
temp = getMeanStdCountRow(temp,true);
tempHeader = [num2cell(G.lapNum);statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Information Per Spike'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.infoPerSpike);
temp = getMeanStdCountRow(temp,true);
tempHeader = [num2cell(G.lapNum);statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Shift Relative to Prior Lap (in bins)'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.shiftRPL);
temp = getMeanStdCountRow(temp,true);
tempHeader = [RowHeader;statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Shift Relative to Original Position (in bins)'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.shiftROP);
temp = getMeanStdCountRow(temp,true);
tempHeader = [RowHeader;statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Absolute Shift (in bins)'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.shiftAbs);
temp = getMeanStdCountRow(temp,true);
tempHeader = [RowHeader;statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Average Correlation Within 5-Lap Sets'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.rAvg5L);
temp = getMeanStdCountRow(temp,true);
tempHeader = [num2cell(G.lapNum);statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Average Correlation Across All Laps'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.rAvgAll);
temp = getMeanStdCountRow(temp,true);
tempHeader = [num2cell(G.lapNum);statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

sheet(end+1,1) = {'Correlation of Each Lap With the Next'};
sheet(end+1,2:end) = [G.cellName(G.cellSel)',statsHeader'];
temp = getMeanStdCountCol(O.rRPL);
temp = getMeanStdCountRow(temp,true);
tempHeader = [RowHeader;statsHeader];
temp = [tempHeader,temp];
[M,N] = size(temp);
sheet(end+1:length(sheet(:,1))+M,1:N) = temp;
sheet(end+1,:) = {[]};

end
% -------------------------------------------------------------------------


% Get Mean, Standard Deviation, and Count Per Column
% -------------------------------------------------------------------------
function [temp] = getMeanStdCountCol(temp)
% get size of input array
sizeTemp = size(temp);
numRows = sizeTemp(1);
numCols = sizeTemp(2);
t = temp;
% find any entries 
[r,c] = find(isnan(temp));
rowCol = [r,c];
numRowsNaN = zeros(1,numCols)/0;
for m = 1:numCols
    numRowsNaN(m) = length(find(rowCol(:,2) == m));
end
rowsAct = bsxfun(@minus,numRows*ones(1,numCols),numRowsNaN);
t(isnan(t)) = 0;
meanTemp = sum(t,1)./rowsAct;
varTemp = (bsxfun(@minus,temp,meanTemp)).^2;
varTemp(isnan(varTemp)) = 0;
varTemp = sum(varTemp,1)./(rowsAct-1);
stdTemp = varTemp.^(1/2);

temp = [temp;meanTemp;stdTemp;rowsAct];
end
% -------------------------------------------------------------------------


% Get Mean, Standard Deviation, and Count Per Row
% -------------------------------------------------------------------------
function [temp] = getMeanStdCountRow(temp,clear3by3Option)
sizeTemp = size(temp);
numRows = sizeTemp(1);
numCols = sizeTemp(2);
t = temp;
[r,c] = find(isnan(temp));
rowCol = [r,c];
numColsNaN = zeros(numRows,1)/0;
for m = 1:numRows
    numColsNaN(m) = length(find(rowCol(:,1) == m));
end
colsAct = bsxfun(@minus,numCols*ones(numRows,1),numColsNaN);
t(isnan(t)) = 0;
meanTemp = sum(t,2)./colsAct;

varTemp = (bsxfun(@minus,temp,meanTemp)).^2;
varTemp(isnan(varTemp)) = 0;
varTemp = sum(varTemp,2)./(colsAct-1);
stdTemp = varTemp.^(1/2);

temp = [temp,meanTemp,stdTemp,colsAct];
if clear3by3Option, temp(end-2:end,end-2:end) = 0/0; end
temp = num2cell(temp);
end
% -------------------------------------------------------------------------


% Import Data Tool Data
% -------------------------------------------------------------------------
function [MT,G] = importDataToolData(UI,G,MT)
fileName = UI.name;
path     = UI.path;

excelObj = actxserver('Excel.Application');
for n = 1:G.numFiles;
    excelWorkbook = excelObj.workbooks.Open([path filesep fileName{n}]);
    worksheets = excelObj.sheets;
    numSheets = worksheets.Count;
    SheetIndex = 1;
    m = 0;
    while SheetIndex <= numSheets
        SheetName = worksheets.Item(SheetIndex).Name;
        if strncmp(SheetName,'Sheet',5) == 0 && strcmp(SheetName,'TS & Bin') == 0
            m = m + 1;
            WorkSheet = excelWorkbook.Sheets.Item(SheetIndex);
            invoke(WorkSheet,'Activate');
            DataRange = excelObj.ActiveSheet.UsedRange;
            temp = cell2mat(DataRange.Value);
            if n == 1
                MT{m,1}.name = SheetName;
                G.cellName{m,1} = SheetName;
            end
            if n == 1 && m == 1, G.bins      = temp(:,1); end
            if m == 1,           G.time(:,n) = temp(:,2); end
            MT{m,1}.freq(:,n) = temp(:,4);
            MT{m,1}.SC(:,n) = temp(:,3);
            SheetIndex = SheetIndex + 1;
        else
            SheetIndex = SheetIndex + 1;
        end
    end
    excelWorkbook.Close;
end
excelObj.Quit; delete(excelObj);
end
% -------------------------------------------------------------------------


% Write Output To File
% -------------------------------------------------------------------------
function [] = writeToExcel(sheet)

[Name,Path] = uiputfile('*.xls','Save As');
excelObj = actxserver('Excel.Application');
Workbook = excelObj.Workbooks.Add;
numSheets = length(sheet);

for n = numSheets:-1:1
    Size = size(sheet{n}.data);
    OverShoot = Size(2)-26;
    if OverShoot > 0
        NumAlphabets = ceil(Size(2)/26)-1;
        ExcelRange = ['A1:' char('a'+NumAlphabets-1) char('a'+OverShoot-26*(NumAlphabets-1)-1) num2str(Size(1))];
    else
        ExcelRange = ['A1:' char('a'+Size(2)-1) num2str(Size(1))];
    end
        
    Worksheet{n} = excelObj.sheets.Add;
    Worksheet{n}.Name = sheet{n}.name;
    Range{n} = get(Worksheet{n},'Range',ExcelRange);
    Range{n}.Value = sheet{n}.data;
end

% Delete Default Worksheets
Worksheets = excelObj.sheets;
numSheets = Worksheets.Count;
sheetIdx = 1;
sheetIdx2 = 1;
while sheetIdx2 <= numSheets
   sheetName = Worksheets.Item(sheetIdx).Name(1:end-1);
   if strcmp(sheetName,'Sheet')
      Worksheets.Item(sheetIdx).Delete;
   else
      sheetIdx = sheetIdx + 1;
   end
   sheetIdx2 = sheetIdx2 + 1;
end

Workbook.SaveAs([Path Name]); Workbook.Close;
excelObj.Quit; delete(excelObj);

end
% -------------------------------------------------------------------------




















