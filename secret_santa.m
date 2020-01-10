function secret_santa
% input xls file with list of names
% first column (A) is the names of people participating
% second column (B) is their email adress
% requires outlook 

[secret_file, secret_path]= uigetfile('*.xlsx');
[~,secret_names,~]= xlsread([secret_path secret_file]);

%randomise
random_order= randperm(size(secret_names,1));
new_order_names= secret_names(random_order,1);

% write to excel file (second sheet)
column_names= [{'Person..'} {'Gives to:'} {'Email Address'}];
% a pair cannot have reciprocal gifts !
attribute= [new_order_names circshift(new_order_names,-1) secret_names(random_order,2)];
xlswrite([secret_path secret_file],[column_names;attribute],2);

% send emails
for this_person=1:size(attribute,1)
    % Create object and set parameters.
    h = actxserver('outlook.Application');
    mail = h.CreateItem('olMail');
    mail.Subject = 'Secret Santa';
    mail.To = char(attribute(this_person,3));
    mail.BodyFormat = 'olFormatHTML';
    task_instr= strjoin(['Ho Ho Ho! <br/> <br/> This year you have been randomly selected to give a gift (of £10 max) to...  <br/> <br/>' attribute(this_person,2) '<br/> <br/> Good luck! <br/> <br/> PS: This is a computer-generated email, please do not reply <br/> <br/>']);
    mail.HTMLBody = task_instr;
    
    % Send message and release object.
    mail.Send;
    h.release;
end

end
