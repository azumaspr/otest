clear;
close all;
yl=zeros(4,4);
y=zeros(4,4);

DataY1=readmatrix('Admittance_data_posian.xlsx')

%ノードアドミタンス行列の作成
for i=1:4
    for ii=1:4
        if i==ii
            yl(i,ii)=0;
        else
            for ia=1:5
                a=DataY1(ia,1);
                b=DataY1(ia,2);
                if a==i && b==ii
                    yl(i,ii)=complex(DataY1(ia,3),DataY1(ia,4));
                elseif a==ii && b==i
                    yl(i,ii)=complex(-DataY1(ia,3),-DataY1(ia,4));
                end
            end
        end
    end
end

for k=1:4
    for kk=1:4
        if k==kk
            for m=1:4
                y(k,kk)=y(k,kk)+yl(k,m);
            end
        else
            y(k,kk)=yl(k,kk);
        end
    end
end
y