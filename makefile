all :
	echo "see makefile"
	echo "to edit in local, see https://inaenomaki.hatenablog.com/entry/2020/09/23/004146"

pull:
	clasp pull

push:
	clang-format-13 -i コード.js
	git add -A . && git commit -m "clasp pushed on `date +%Y.%m.%d.%H.%M.%S`" && git push origin master &
	clasp push

login:
	clasp login

install:
	# wslのubuntuを想定
	# clang-formatのインストール
	sudo apt update
	sudo apt install clang-format-13
	# nvmのインストール
	curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.35.2/install.sh | bash
	nvm install --lts
	# v14.18.1 が入った想定
	nvm use v14.18.1
	npm install -g @google/clasp

