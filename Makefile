
.PHONY:build

build:
	sudo pyinstaller --onefile --windowed --icon="common/Nomina.ico" --add-data "common/deps/tabula-1.0.5-jar-with-dependencies.jar;tabula" --add-data "common/Nomina.ico;." --name "Nomina" main.py