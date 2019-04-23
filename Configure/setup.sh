if [ `yum repolist|grep epel|wc -l` -eq 0 ]

then

sudo yum install https://dl.fedoraproject.org/pub/epel/epel-release-latest-7.noarch.rpm -y 

sudo yum-config-manager --enable epel

sudo yum update -y

fi

sudo yum install gcc -y




tar -xzf ncurses-6.1.tar.gz

cd ncurses-6.1

./configure

sudo make

sudo make install

cd ..

rm -rf ncurses-6.1




tar -xzf mutt-1.11.4.tar.gz

cd mutt-1.11.4

./configure

sudo make

sudo make install

cd ..

rm -rf mutt-1.11.4



sudo yum install python2-pip -y

sudo pip install --upgrade pip


sudo pip install -r requirements.txt

sudo chmod 755 ../*.sh

which mutt

which pip

read
