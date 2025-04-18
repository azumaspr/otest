%FILE name: pocian.m
%Power circuit analysis using Jacobian method
clear;
close all;
y=zeros(4,4);
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
%ノードアドミタンス行列の作成
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
y;
%母線電圧の初期値の設定
itn=0;
erv=1.0E-5;%収束判定のε
vv=v;
cu=zeros(3,1);
vdmax=1.0E+3;%近似解の誤差の初期範囲
while vdmax>erv
    itn=itn+1;
    for ii=1:3
        cu(ii)=0.0;
        for j=1:4
            if ii==j
                cu(ii)=cu(ii)+ci(ii);
            else
                cu(ii)=cu(ii)-y(ii,j)*vv(j);
            end
        end
        v(ii)=cu(ii)/y(ii,ii);
    end
    for l=1:3
        avd(l)=abs(v(l)-vv(l));
    end
    vdmax=max(avd);
    itn;%反復回数の表示
    vdmax;%反復回数itnのときの近似解の表示
    %複素数ベクトルvの実部と虚部を分けた実数行列out_vを作りコマンドウィンドに表示
    out_v=[real(v),imag(v)];
    vv=v;
end
%反復計算で求めた複素電圧vをアドミタンス行列yに掛け算して注入電流liと一致するのかをチェック
check_ci=y*v;
%エクセルファイルOutput_data_posian.xlsxにitn,vdmax,out_vを書き込む
writematrix(itn,'Output_data_posian.xlsx','Range','A1')
writematrix(vdmax,'Output_data_posian.xlsx','Range','A2')
writematrix(out_v,'Output_data_posian.xlsx','Range','A3')