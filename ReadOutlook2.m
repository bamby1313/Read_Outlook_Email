function [Email]= ReadOutlook2(options)
%  Scraping emails from Microsoft Outlook
%  Functionality which imports readed, unreaded emails from inbox or
%  or outlook folders and subfolders
%  Extracts their subjects, bodies and can save their attachements
%
% Function Inputs
%    Basic import functionality Varargin
%    ------------------------------
%       AccountName = outlook account name
%       Folder      = outlook folder name
%       SubFolder   = outlook subfolder name
%       SavePath    = path to save the attachments
%       Read        = 1,  reads only the UnRead emails
%       Mark        = 1,  marks UnRead emails as read
%       
%
% Author: Pierre Harouimi,
%		pharouim@mathworks.com
%--------------------------------------------------------------------------
% Examples:
%
% %     Reads all emails from your default inbox
%          mails = ReadOutlook2
%
% %     Reads all Unread emails from your inbox
%          mails = ReadOutlook2("Read", 1)
%
% %     Reads all Unread emails from your inbox and mark them as read
%          mails = ReadOutlook2("Read", 1, "Mark", 1)

arguments
    options.AccountName string = ""
    options.Folder      string  = ""
    options.SubFolder   string  = ""
    options.SavePath    {mustBeFolder}
    options.Read        logical = false
    options.Mark        logical = false
end

%% Connects to Outlook
outlook = actxserver('Outlook.Application');
mapi    = outlook.GetNamespace('mapi');
if options.AccountName == ""
    INBOX = mapi.GetDefaultFolder(6);
else
    account_numbers = mapi.Folders.Count;
    for i = 1 : account_numbers
        name = mapi.Folders.Item(i).Name;
        if name == options.AccountName
            INBOX = mapi.Folders.Item(i);
            break
        end
    end
end

%% Retrieving UnRead or read emails / save or not save attachments
if options.Folder == "" && options.SubFolder == ""
    % reads Inbox only
    count = INBOX.Item.Count;
    Email = cell(count,2);
elseif options.Folder ~= ""
    % reads Inbox folder
    folder_numbers = INBOX.Folders.Count;
    % find folder / subfolder's outlookindex
    for i = 1:folder_numbers
        name = INBOX.Folders(1).Item(i).Name;
        if name == options.Folder
            idx = i;
        end
    end
    switch options.SubFolder
        % working for folder emails
        case ''
            % number of emails
            count = INBOX.Folders(1).Item(idx).Items.Count;
            % cell for emailbody
            Email = cell(count,2);
        otherwise
            % Search for nth Inbox folder and count sub-folders
            folder_numbers = INBOX.Folders(1).Item(idx).Folders(1).Count;
            % find Outlook Subfolder Index
            for i=1:folder_numbers
                name = INBOX.Folders(1).Item(idx).Folders(1).Item(i).Name;
                if name == options.SubFolder
                    s= i;
                end
            end
            % number of emails
            count = INBOX.Folders(1).Item(idx).Folders(1).Item(s).Items.Count;
            % cell for emailbody
            Email = cell(count,2);
    end
end

%% download & read emails
for i = 1:count
    if options.Read == 1 % only unreads emails
        % inbox
        if options.Folder == "" && options.SubFolder == ""
            UnRead = INBOX.Items.Item(count+1-i).UnRead;
            % folder
        elseif options.Folder ~= "" && options.SubFolder == ""
            UnRead = INBOX.Folders(1).Item(idx).Items(1).Item(count+1-i).UnRead;
            % subfolder
        elseif options.Folder ~= "" && options.SubFolder ~= ""
            UnRead = INBOX.Folders(1).Item(idx).Folders(1).Item(s).Item(1).Item(count+1-i).UnRead;
        end

        if UnRead
            % inbox
            if options.Folder == "" && options.SubFolder == ""
                if options.Mark == 1
                    INBOX.Items.Item(count+1-i).UnRead=0;
                end
                email = INBOX.Items.Item(count+1-i);
                % folder
            elseif options.Folder ~= "" && options.SubFolder == ""
                if options.Mark == 1
                    INBOX.Folders(1).Item(idx).Items(1).Item(count+1-i).UnRead=0;
                end
                email = INBOX.Folders(1).Item(idx).Items(1).Item(count+1-i);
                % subfolder
            elseif options.Folder ~= "" && options.SubFolder ~= ""
                if options.Mark == 1
                    INBOX.Folders(1).Item(idx).Folders(1).Item(s).Item(1).Item(count+1-i).UnRead=0;
                end
                email = INBOX.Folders(1).Item(idx).Folders(1).Item(s).Items.Item(count+1-i);
            end
        end
    else   % all emails
        % inbox
        if options.Folder == "" && options.SubFolder == ""
            email = INBOX.Items.Item(count+1-i);
            % folder
        elseif options.Folder ~= "" && options.SubFolder == ""
            email = INBOX.Folders(1).Item(idx).Items(1).Item(count+1-i);
            % subfolder
        elseif options.Folder ~= "" && options.SubFolder ~= ""
            email = INBOX.Folders(1).Item(idx).Folders(1).Item(s).Items.Item(count+1-i);
        end
        UnRead = 1; %pseudo for next step
    end
    if UnRead
        % read and save body
        subject     = email.get('Subject');
        body        = email.get('Body');
        Email{i,1}  = subject;
        Email{i,2}  = body;
        if options.SavePath ~= ""
            attachments = email.get('Attachments');
            if attachments.Count >= 1
                fname   = "Video" + extract(string(subject), digitsPattern);
                full    = fullfile(options.SavePath, fname);
                attachments.Item(1).SaveAsFile(full)
            end
        end
    end
end

Email(all(cellfun('isempty', Email),2),:)=[];

end

