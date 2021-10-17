all :
	echo "see makefile"
	echo "to edit in local, see https://inaenomaki.hatenablog.com/entry/2020/09/23/004146"

pull:
	clasp pull

push:
	clang-format-13 -i コード.js
	clasp push

login:
	clasp login

