# Mladen Papulic
# 25.11.2016.
# Script for creating THD SIDK tests.


import xlrd
import os
from config import TEST_NAME
from config import GENERATE_CFG
from config import MAT_CBR_ON
from config import MAT_CBR_OFF
from config import MAT_CBR_OFF_MLP
from config import PATH_TO_OUTPUT
from config import RENAME_DUT_TO_REF
from config import MERGE
from config import MOVE_DUT
from config import DEL_STREAM
from config import ASIO_INPUT_MAT
from config import ASIO_INPUT_NOMAT
from config import LOAD_1FS_LEGACY_0
from config import LOAD_2FS_LEGACY_0
from config import LOAD_4FS_LEGACY_0
from config import LOAD_1FS_LEGACY_1
from config import LOAD_2FS_LEGACY_1
from config import LOAD_4FS_LEGACY_1
from config import ASIO_4FS_LEGACY_1
from config import ASIO_2FS_LEGACY_1
from config import ASIO_ELSE
from config import START_PLAYING

# Open a workbook:

book = xlrd.open_workbook("TrueHD_SIDK.xls")
worksheet = book.sheet_by_index(0)


def generator():
    if __name__ == "__main__":
        x = 1

        while x < 1386:


        # Set variables..

            test_name = worksheet.cell_value(rowx=x, colx=0)
            stream_path = worksheet.cell_value(rowx=x, colx=1)
            test_folder = worksheet.cell_value(rowx=x, colx=2)
            out_file_name = worksheet.cell_value(rowx=x, colx=3)
            test_path = worksheet.cell_value(rowx=x, colx=4)
            cmd_params = worksheet.cell_value(rowx=x, colx=5)
            legacy = worksheet.cell_value(rowx=x, colx=6)
            fs = worksheet.cell_value(rowx=x, colx=7)
            mat_cbr = worksheet.cell_value(rowx=x, colx=8)
            y = stream_path.rfind('\\')
            y = y + 1
            play_file = stream_path[y:]



        # Create mask..

            os.system("duet_lib_create_mask.py" + " " + '"' + cmd_params + '"')
            mask = open('asio_masks.txt','r')
            mask = mask.read()

        # Make directory..

            dir = 'tests_new' + "\\" + test_path
            if not os.path.exists(dir):
                os.makedirs(dir)

        # Open and write file..

            file = open(os.path.join(dir, test_name + ".tst"), "w")

            file.write(TEST_NAME.format(x=x, test_name=test_name))

            file.write(GENERATE_CFG.format(cmd_params=cmd_params,
                                               stream_path=stream_path[:-4], test_path=test_path))

            if mat_cbr == 1 and stream_path[-3:] == "mat":
                file.write(MAT_CBR_ON.format(stream_path=stream_path, play_file=play_file[:-4]))

            elif mat_cbr == 0 and stream_path[-3:] == "mat":
                file.write(MAT_CBR_OFF.format(stream_path=stream_path, play_file=play_file[:-4]))

            elif mat_cbr == 0 and stream_path[-3:] != "mat":
                file.write(MAT_CBR_OFF_MLP.format(stream_path=stream_path, play_file=play_file[:-4]))

            if mat_cbr == 1:
                file.write(ASIO_INPUT_MAT.format(play_file=play_file[:-4]))

            else:
                file.write(ASIO_INPUT_NOMAT.format(play_file=play_file[:-4]))

            if fs == "1fs" and legacy == 0:
                file.write(LOAD_1FS_LEGACY_0)

            elif fs == "2fs" and legacy == 0:
                file.write(LOAD_2FS_LEGACY_0)

            elif fs == "4fs" and legacy == 0:
                file.write(LOAD_4FS_LEGACY_0)

            elif fs == "1fs" and legacy == 1:
                file.write(LOAD_1FS_LEGACY_1)

            elif fs == "2fs" and legacy == 1:
                file.write(LOAD_2FS_LEGACY_1)

            elif fs == "4fs" and legacy == 1:
                file.write(LOAD_4FS_LEGACY_1)

            file.write(PATH_TO_OUTPUT)

            if fs == "4fs" and legacy == 1:
                file.write(ASIO_4FS_LEGACY_1.format(mask=mask))

            elif fs == "2fs" and legacy == 1:
                file.write(ASIO_2FS_LEGACY_1.format(mask=mask))

            else:
                file.write(ASIO_ELSE.format(mask=mask))

            file.write(START_PLAYING)

            file.write(RENAME_DUT_TO_REF.format(test_folder=test_folder))
            file.write(MERGE.format(test_folder=test_folder))
            file.write(MOVE_DUT.format(test_folder=test_folder, out_file_name=out_file_name))
            file.write(DEL_STREAM.format(play_file=play_file[:-4]))

            print str(x) + " " + "Test" + " " + '"' + test_name + '"' + " " + "created!"
            x += 1
            file.close()

        print "Done! " + str(x-1) + " tests are created!"
    else:
        print "__name__ is "+__name__

generator()


