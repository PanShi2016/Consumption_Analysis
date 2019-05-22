% process xls files

% labels for each year
years = {'2019'};

% labels for each month
months = {'01';'02';'03';'04';'05';'06';'07';'08';'09';'10';'11';'12'};
months_num = {'1';'2';'3';'4';'5';'6';'7';'8';'9';'10';'11';'12'};
days = {'31';'Feb';'31';'30';'31';'30';'31';'31';'30';'31';'30';'31'};

% labels for consumption patterns
% labels_canteen = {'西一';'喻园';'集锦园';'东一';'学一';'东教工';'东园';'韵苑风味'}; % labels for canteen
labels_canteen = {'西一';'喻园';'集锦园';'东一';'学一';'学二';'东教工';'东园';'韵苑风味';'百品屋';'百景';'集贤楼'}; % labels for canteen
labels_supmark = {'东学超市';'教工超市';'韵苑超市'}; % labels for supermarket
labels_sports = {'主校区体育场馆'}; % labels for sports
labels_others = {'博士生公寓';'自助售货机';'校医院';'充电桩'}; % labels for others

consume_months = [];
consume_months_can = [];
consume_canteen = zeros(length(labels_canteen),length(months));
Alltime = [];
dinner_consume = [];

for y = 1 : length(years)
    if (mod(str2num(years{y}),4)==0 && mod(str2num(years{y}),100)~=0) || (mod(str2num(years{y}),400)==0)
        days{2} = '29';
    else
        days{2} = '28';
    end
    for i = 1 : length(months)
        % load xls files
        [num,txt,~] = xlsread([years{y},months{i},'.xls']);
        num_len = length(num);
        txt_len = length(txt);
        
        % total consumption
        index = find(num(:,1)<0);
        consume_all = abs(sum(num(index,1)));
        consume_months = [consume_months,consume_all];
        
        % consumption for each pattern
        can_ind = [];
        can_label = [];
        can_time = [];
        sup_ind = [];
        spo_ind = [];
        oth_ind = [];
        
        for j = 2 : txt_len
            % statistics for canteen
            can_vec = regexp(txt{j,9},labels_canteen);
            can_id = find(cellfun(@(x)~isempty(x),can_vec));
            if ~isempty(can_id)
                can_ind = [can_ind,j-1];
                can_label = [can_label,can_id];
            end
            % statistics for supermarket
            sup_vec = regexp(txt{j,9},labels_supmark);
            sup_id = find(cellfun(@(x)~isempty(x),sup_vec));
            if ~isempty(sup_id)
                sup_ind = [sup_ind,j-1];
            end
            % statistics for sports
            spo_vec = regexp(txt{j,9},labels_sports);
            spo_id = find(cellfun(@(x)~isempty(x),spo_vec));
            if ~isempty(spo_id)
                spo_ind = [spo_ind,j-1];
            end
            % statistics for others
            oth_vec = regexp(txt{j,9},labels_others);
            oth_id = find(cellfun(@(x)~isempty(x),oth_vec));
            if ~isempty(oth_id)
                oth_ind = [oth_ind,j-1];
            end
        end
        
        % consumption for canteen
        consume_can = abs(sum(num(can_ind,1)));
        consume_months_can = [consume_months_can,consume_can];
        all_time = datevec(txt(index+1,1));
        all_consume = abs(num(index,1));
        can_time = datevec(txt(can_ind+1,1));
        can_consume = abs(num(can_ind,1));
        can_day = can_time(:,3);
        all_day = all_time(:,3);
        num_time = can_time(:,4) + can_time(:,5)/60 + can_time(:,6)/3600;
        am_id = find(num_time > 6.5 & num_time < 10.3);
        noon_id = find(num_time > 10.5 & num_time < 14);
        pm_id = find(num_time > 16 & num_time < 20);
        am_time = num_time(am_id);
        noon_time = num_time(noon_id);
        pm_time = num_time(pm_id);
        Alltime = [Alltime;[mean(am_time),mean(noon_time),mean(pm_time),std(am_time),std(noon_time),std(pm_time)]];
        if (sum(num_time) - sum(am_time) - sum(noon_time) -sum(pm_time)) < 1
            % fprintf('Time normal\n');
            fprintf('时间正常\n');
        else
            % fprintf('Time anomaly\n');
            fprintf('时间异常\n');
        end
        
        am_consume = sum(can_consume(am_id));
        noon_consume = sum(can_consume(noon_id));
        pm_consume = sum(can_consume(pm_id));
        dinner_consume = [dinner_consume;[am_consume,noon_consume,pm_consume]];
        
        day_len = str2num(days{i});
        consume_day = zeros(1,day_len);
        consume_day_can = zeros(1,day_len);
        time_day = zeros(day_len,3);
        dinner_consume_day = zeros(day_len,3);
        for s = 1 : day_len
            all_day_id = find(all_day == s);
            day_id = find(can_day == s);
            if ~isempty(all_day_id)
                consume_day(1,s) = sum(all_consume(all_day_id));
            end
            if ~isempty(day_id)
                consume_day_can(1,s) = sum(can_consume(day_id));
                am_day_id = intersect(am_id,day_id);
                noon_day_id = intersect(noon_id,day_id);
                pm_day_id = intersect(pm_id,day_id);
                if ~isempty(am_day_id)
                    time_day(s,1) = mean(num_time(am_day_id));
                    dinner_consume_day(s,1) = sum(can_consume(am_day_id));
                end
                if ~isempty(noon_day_id)
                    time_day(s,2) = mean(num_time(noon_day_id));
                    dinner_consume_day(s,2) = sum(can_consume(noon_day_id));
                end
                if ~isempty(pm_day_id)
                    time_day(s,3) = mean(num_time(pm_day_id));
                    dinner_consume_day(s,3) = sum(can_consume(pm_day_id));
                end
            end
        end
        
        % plot figure for all consumption and canteen consumption of each day
        figure;
        data = [consume_day_can;consume_day]';
        maxdata = max(max(data));
        upper = round(maxdata/10 + 2)*10;
        colour = {[0.3922 0.5843 0.9294],[0 0.749 1]};
        h = zeros(1,2);
        offset = [1 : 0.3 : 2];
        for m = 1 : 2
            h(m) = bar(data(:,m),'FaceColor',colour{m},'BarWidth',0.3);
            hold on;
            set(h(m),'XData',get(h(m),'XData') + offset(m));
        end
        set(gca,'FontSize',7);
        set(gca,'xtick',[2.15:1:(day_len+1.15)]);
        set(gca,'xticklabel',[1:day_len]);
        set(gca,'ytick',[0 : 10 : upper]);
        ylim([0 upper]);
        g = legend('食堂消费','总消费','Location','Northeast');
        set(g,'Orientation','horizon');
        xlabel(['日',' （',years{y},'年',months_num{i},'月','）']);
        ylabel('金额（元）');
        saveas(gcf,[years{y},months{i},'_consume.png']);
        
        % plot figure for meal consumption of each day
        figure;
        data = dinner_consume_day;
        maxdata = max(max(data));
        upper = round(maxdata/10 + 1)*10;
        colour = {[0.5294 0.8078 0.9804],[0.3922 0.5843 0.9294],[0 0.749 1]};
        h = zeros(1,3);
        offset = [1 : 0.2 : 2];
        for m = 1 : 3
            h(m) = bar(data(:,m),'FaceColor',colour{m},'BarWidth',0.2);
            hold on;
            set(h(m),'XData',get(h(m),'XData') + offset(m));
        end
        set(gca,'FontSize',7);
        set(gca,'xtick',[2.2:1:(day_len+1.2)]);
        set(gca,'xticklabel',[1:day_len]);
        set(gca,'ytick',[0 : 10 : upper]);
        ylim([0 upper]);
        g = legend('早餐','午餐','晚餐','Location','Northeast');
        set(g,'Orientation','horizon');
        xlabel(['日',' （',years{y},'年',months_num{i},'月','）']);
        ylabel('金额（元）');
        saveas(gcf,[years{y},months{i},'_consume_dinner.png']);
        
        % plot figure for meal time of each day
        figure;
        data = time_day;
        data(find(data==0)) = 6;
        colour = {[0.5294 0.8078 0.9804],[0.3922 0.5843 0.9294],[0 0.749 1]};
        h = zeros(1,3);
        offset = [1 : 0.2 : 2];
        for m = 1 : 3
            h(m) = bar(data(:,m),'FaceColor',colour{m},'BarWidth',0.2);
            hold on;
            set(h(m),'XData',get(h(m),'XData') + offset(m));
        end
        set(gca,'FontSize',7);
        set(gca,'xtick',[2.2:1:(day_len+1.2)]);
        set(gca,'xticklabel',[1:day_len]);
        % set(gca,'ytick',[0 : 2 : 22]);
        % ylim([0 22]);
        set(gca,'ytick',[6:2:22]);
        ylim([6 22]);
        g = legend('早餐','午餐','晚餐','Location','Northeast');
        set(g,'Orientation','horizon');
        xlabel(['日',' （',years{y},'年',months_num{i},'月','）']);
        ylabel('时间（时钟）');
        saveas(gcf,[years{y},months{i},'_time_day.png']);
        
        % consumption for different canteens
        for k = 1 : length(labels_canteen)
            label_id = find(can_label == k);
            if ~isempty(label_id)
                consume_canteen(k,i) = consume_canteen(k,i) + abs(sum(num(can_ind(label_id),1)));
            end
        end
        
        % consumption for supermarket
        consume_sup = abs(sum(num(sup_ind,1)));
        
        % consumption for sports
        consume_spo = abs(sum(num(spo_ind,1)));
        
        % consumption for others
        consume_oth = abs(sum(num(oth_ind,1)));
        
        consume_sum = consume_can + consume_sup + consume_spo + consume_oth;
        
        if abs(consume_all - consume_sum) < 0.01
            % fprintf('Consumption normal\n');
            fprintf('消费正常\n');
        else
            % fprintf('Consumption anomaly\n');
            fprintf('消费异常\n');
        end
    end
    
    % plot figure for all consumption and canteen consumption of 2018
    figure;
    data = [sum(consume_months_can),sum(consume_months)];
    colour = {[0.3922 0.5843 0.9294],[0 0.749 1]};
    h = zeros(1,2);
    offset = [-0.65 : 0.2 : 1];
    for i = 1 : 2
        h(i) = bar(data(i),'FaceColor',colour{i},'BarWidth',0.2);
        hold on;
        set(h(i),'XData',get(h(i),'XData') + offset(i));
    end
    set(gca,'FontSize',15);
    set(gca,'xtick',[0.45:1:2]);
    set(gca,'xticklabel',years{y});
    set(gca,'ytick',[0 : 2000 : 10000]);
    ylim([0 10000]);
    g = legend('食堂消费','总消费','Location','Northeast');
    set(g,'Orientation','horizon');
    xlabel('年份');
    ylabel('金额（元）');
    saveas(gcf,[years{y},'_consume.png']);
    
    % plot figure for all consumption and canteen consumption of each month
    figure;
    data = [consume_months_can;consume_months]';
    colour = {[0.3922 0.5843 0.9294],[0 0.749 1]};
    h = zeros(1,2);
    offset = [-0.65 : 0.3 : 1];
    for i = 1 : 2
        h(i) = bar(data(:,i),'FaceColor',colour{i},'BarWidth',0.3);
        hold on;
        set(h(i),'XData',get(h(i),'XData') + offset(i));
    end
    set(gca,'FontSize',15);
    set(gca,'xtick',[0.5:1:12]);
    set(gca,'xticklabel',months_num);
    set(gca,'ytick',[0 : 200 : 1000]);
    ylim([0 1000]);
    g = legend('食堂消费','总消费','Location','Northeast');
    set(g,'Orientation','horizon');
    xlabel(['月份（',years{y},'年）']);
    ylabel('金额（元）');
    saveas(gcf,[years{y},'_consume_month.png']);
    
    % plot figure for different canteen consumption
    figure;
    data = consume_canteen';
    offset = [1 : 1.2 : 24];
    for i = 1 : size(data,1)
        h(i,1:size(data,2)) = bar([data(i,:);[1:size(data,2)]],'stacked');
        hold on;
        set(h(i,:),'XData',offset(i));
    end
    % Set colors
    colour = {[0.2 0.6 0.6],[0.2 0.4 0.7],[0.4 0.5 0.8],[0.6 0.8 0.6],...
        [0.1 0.6 1],[0.3 0.6 0.8],[0.4 0.7 1],[0.3 0.5 0.7],[0.7 0.8 0.9]};
    for j = 1 : size(data,2)
        set(h(:,j),'facecolor',colour{j});
    end
    set(gca,'FontSize',15);
    set(gca,'box','on','xtick',[1: 1.2 : 24],...
        'xticklabels',months_num);
    set(gca,'ytick',[0 : 200 : 1000]);
    ylim([0 1000]);
    g = legend(labels_canteen,'Location','bestoutside');
    xlabel(['月份（',years{y},'年）']);
    ylabel('累计金额（元）');
    saveas(gcf,[years{y},'_consume_canteen.png']);
    
    % plot figure for meal time of each month
    figure;
    data = Alltime;
    colour = {[0.5294 0.8078 0.9804],[0.3922 0.5843 0.9294],[0 0.749 1]};
    h = zeros(1,3);
    offset = [-0.7 : 0.2 : 1];
    for i = 1 : 3
        h(i) = bar(data(:,i),'FaceColor',colour{i},'BarWidth',0.2);
        hold on;
        set(h(i),'XData',get(h(i),'XData') + offset(i));
    end
    x = [0.3:1:11.3,0.5:1:11.5,0.7:1:11.7];
    e = errorbar(x,[data(:,1)',data(:,2)',data(:,3)'],[data(:,4)',data(:,5)',data(:,6)'],'LineStyle','none');
    e.Color = 'black';
    set(gca,'FontSize',15);
    set(gca,'xtick',[0.5:1:12]);
    set(gca,'xticklabel',months_num);
    set(gca,'ytick',[7:2:11,12,13:2:17,18,19:2:21]);
    ylim([7 21]);
    g = legend('早餐','午餐','晚餐','Location','Northeast');
    set(g,'Orientation','horizon');
    xlabel(['月份（',years{y},'年）']);
    ylabel('时间（时钟）');
    saveas(gcf,[years{y},'_time_month.png']);
    
    % plot figure for meal consumption of each month
    figure;
    data = dinner_consume;
    colour = {[0.5294 0.8078 0.9804],[0.3922 0.5843 0.9294],[0 0.749 1]};
    h = zeros(1,3);
    offset = [-0.7 : 0.2 : 1];
    for i = 1 : 3
        h(i) = bar(data(:,i),'FaceColor',colour{i},'BarWidth',0.2);
        hold on;
        set(h(i),'XData',get(h(i),'XData') + offset(i));
    end
    set(gca,'FontSize',15);
    set(gca,'xtick',[0.5:1:12]);
    set(gca,'xticklabel',months_num);
    set(gca,'ytick',[0:100:500]);
    ylim([0 500]);
    g = legend('早餐','午餐','晚餐','Location','Northeast');
    set(g,'Orientation','horizon');
    xlabel(['月份（',years{y},'年）']);
    ylabel('金额（元）');
    saveas(gcf,[years{y},'_consume_dinner.png']);
end