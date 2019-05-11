all: build

build: clean copy package

clean:
	rm -rf dist/ && mkdir dist/

copy:
	cp manifest.json dist/

package:
	cd dist && zip -r app.zip . --exclude=./app.zip

