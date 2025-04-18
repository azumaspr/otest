%FILE name: pocian.m
%Power circuit analysis using Jacobian method
clear;
close all;
y=zeros(4,4);
ys=zeros(3,3);
y4=zeros(3,1);
ci=zeros(3,1);
v=ones(4,1);
%各送電線のアドミタンス
yl12=5.0-12.0i;
yl13=5.0-12.0i;
yl14=10.0-40.0i;
yl23=10.0-30.0i;
yl34=-20.0i;
%母線＃４の電圧
v(4)=1.05;
%注入電流Iiデータファイルからの取り込み
DataIn=readmatrix('Input_data_pocian.xlsx');
DataIn;%DataInの中身をコマンドウィンドウに表示して確認
for ii=1:3
    ci(ii)=complex(DataIn(ii,1),DataIn(ii,2));
end
ci;
v;
%ノードアドミタンス行列(とその一部)の作成
y(1,1)=yl12+yl13+yl14;
y(1,2)=-yl12;
y(1,3)=-yl13;
y(1,4)=-yl14;
y(2,1)=-yl12;
y(2,2)=yl12+yl23;
y(2,3)=-yl23;
y(2,4)=0;
y(3,1)=-yl13;
y(3,2)=-yl23;
y(3,3)=yl13+yl23+yl34;
y(3,4)=-yl34;
y(4,1)=-yl14;
y(4,2)=0;
y(4,3)=-yl34;
y(4,4)=yl14+yl34;
y
for ii=1:3
    for iii=1:3
        ys(ii,iii)=y(ii,iii);
    end
    y4(ii,1)=y(ii,1);
end
ys
y4
%母線電圧を求める計算
i=0;
cu=zeros(3,1);
cu=ys\(ci-y4*v(4));
for i=1:3
    v(i)=cu(i);
end
out_v=[real(v),imag(v)];
%反復計算で求めた複素電圧vをアドミタンス行列yに掛け算して注入電流liと一致するのかをチェック
check_ci=y*v;
%エクセルファイルOutput_data_posian.xlsxにout_vを書き込む
writematrix(out_v,'Output_data_posian.xlsx','Range','A1')