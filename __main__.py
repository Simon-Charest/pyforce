from glob import glob
from io import BufferedReader, BytesIO
from itertools import product
from msoffcrypto import OfficeFile
from msoffcrypto.format.ooxml import OOXMLFile
from sys import argv


def main(verbose: bool = False) -> None:
    files: list[str] = glob("data/**/*.xls", recursive=True) + glob("data/**/*.xlsx", recursive=True)
    file: str
    attempt: int = 0

    for file in files:
        characters: tuple[int, ...]
        
        for characters in product(range(65, 91), repeat=6):
            password: str = "".join(chr(character) for character in characters)
            
            if verbose:
                attempt += 1
                print(f"{attempt}: {password}")

            try:
                stream: BufferedReader = open(file, "rb")
                office_file: OOXMLFile = OfficeFile(stream)
                office_file.load_key(password)
                outfile: BytesIO = BytesIO()
                office_file.decrypt(outfile)
                print(password)
                stream.close()
                
                return

            except Exception:
                pass

            finally:
                stream.close()


if __name__ == "__main__":
    main(any(argument in argv for argument in ("-v", "--verbose")))
