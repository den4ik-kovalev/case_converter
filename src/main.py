import os
import sys
from collections import OrderedDict
from enum import Enum, auto
from pathlib import Path

from loguru import logger

from src.library.xls_file import XLSFile
from src.library.yaml_file import YAMLFile


ROOT_DIR = Path(os.path.dirname(sys.argv[0]))  # .exe filepath

INPUT_DIR = ROOT_DIR / "Input"
OUTPUT_DIR = ROOT_DIR / "Output"
LOG_DIR = ROOT_DIR / "Log"
SETTINGS_DIR = ROOT_DIR / "Settings"

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SETTINGS_DIR, exist_ok=True)

CONFIG_PATH = SETTINGS_DIR / "config.yml"
LOGGER_PATH = LOG_DIR / "error.log"
logger.add(LOGGER_PATH, format="{time} | {level} | {message}", level="ERROR")


class Section(Enum):
    NAME = auto()
    PRECS = auto()
    STEPS = auto()


def convert_txt_to_xlsx(ifp: Path, ofp: Path, config: dict) -> None:

    cfg_start = config["start"]
    cfg_delimiter = config["delimiter"]
    cfg_preconditions = config["preconditions"]
    cfg_steps = config["steps"]
    cfg_bracket = config["bracket"]

    with open(ifp, mode="r", encoding="utf-8") as file:
        lines = file.readlines()

    lines = [line.strip("\n") for line in lines if line]
    lines = [line for line in lines if line]

    if cfg_start:
        start_idx = lines.index(cfg_start)
        lines = lines[(start_idx + 1):]

    current_section = Section.NAME
    current_case = {"name": "", "precs": [], "steps": []}
    cases = []

    for line in lines:
        if line == cfg_delimiter:
            cases.append(current_case)
            current_case = {"name": "", "precs": [], "steps": []}
            current_section = Section.NAME
            continue
        if line == cfg_preconditions:
            current_section = Section.PRECS
            continue
        if line == cfg_steps:
            current_section = Section.STEPS
            continue
        if current_section == Section.NAME:
            name = line.strip()
            current_case["name"] = name
            continue
        if current_section == Section.PRECS:
            prec = line.strip()
            current_case["precs"].append(prec)
            continue
        if current_section == Section.STEPS:
            if cfg_bracket in line:
                step, _, result = line.partition(cfg_bracket)
                step = step.strip()
                result = result.strip()
                result = result[:-1]
                result = result.strip()
            else:
                step = line.strip()
                result = None
            current_case["steps"].append((step, result))

    cases.append(current_case)

    data = []
    for case in cases:

        obj = OrderedDict()
        obj["Наименование"] = case["name"]
        obj["Предусловия"] = None
        obj["Шаги"] = None
        obj["Ожидаемый результат"] = None
        obj["Статус"] = "Готов"
        data.append(obj)

        for prec in case["precs"]:
            obj = OrderedDict()
            obj["Наименование"] = None
            obj["Предусловия"] = prec
            obj["Шаги"] = None
            obj["Ожидаемый результат"] = None
            obj["Статус"] = None
            data.append(obj)

        for step in case["steps"]:
            obj = OrderedDict()
            obj["Наименование"] = None
            obj["Предусловия"] = None
            obj["Шаги"] = step[0]
            obj["Ожидаемый результат"] = step[1]
            obj["Статус"] = None
            data.append(obj)

    output_file = XLSFile(ofp)
    output_file.write(data)

    print(f"{len(cases)} cases -> {ofp}")
    input("Press Enter")


@logger.catch
def main() -> None:

    config = YAMLFile(CONFIG_PATH, auto_create=False).read()
    for file in INPUT_DIR.iterdir():
        output_path = OUTPUT_DIR / (file.stem + ".xlsx")
        convert_txt_to_xlsx(file, output_path, config)


if __name__ == "__main__":
    main()
