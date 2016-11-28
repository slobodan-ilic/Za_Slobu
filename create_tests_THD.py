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

# Open a workbook:

book = xlrd.open_workbook("TrueHD_SIDK.xls")
worksheet = book.sheet_by_index(0)
# TODO: 2. Create a function and call it from main file. Check how the if __name__ == "__main__" idiome works
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
            file.write(
            "////////////////////////////////////////////////////////////////////////////\n"
            "\n"
            "/////////////////////////////////////////////////Set up path to asio input file//////////\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  StringManager_4\n"
            "command = value\n"
            "value = Globals::D_duet_tools\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  StringManager_4\n"
            "command = concat\n"
            )
            if mat_cbr == 1:
                file.write(
                "value =" + " " + play_file[:-4] + ".iec" + "\n"
                )
            else:
                file.write(
                    "value =" + " " + play_file[:-4] + ".mat" + "\n"
                )
            file.write(
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            )
            file.write(
            "////////////////////////////////////////////////////////////\n"
            "// LOAD\n"
            "////////////////////////////////////////////////////////////\n"
                "\n"
            "[step]\n"
            "description =\n"
            "device =  StringManager_6\n"
            "command = value\n"
            "value = Globals::D_duet_InstDir\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device = StringManager_6\n"
            "command = concat\n"
            )
            if fs == "1fs" and legacy == 0:
                file.write(
                "value = Globals::D_duet_load\n"
                )
            elif fs == "2fs" and legacy == 0:
                file.write(
                "value = Globals::D_duet_load_2fs\n"
                )
            elif fs == "4fs" and legacy == 0:
                file.write(
                "value = Globals::D_duet_load_4fs\n"
                )
            elif fs == "1fs" and legacy == 1:
                file.write(
                "value = Globals::D_duet_Legacy_load\n"
                )
            elif fs == "2fs" and legacy == 1:
                file.write(
                "value = Globals::D_duet_Legacy_load2\n"
                )
            elif fs == "4fs" and legacy == 1:
                file.write(
                "value = Globals::D_duet_Legacy_load4\n"
                )
            file.write(
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  CS49834Console\n"
            "command = command\n"
            "value =  StringManager_6::value\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            )
            file.write(PATH_TO_OUTPUT)
            file.write(
            "/////////////////////////////////////////////////////\n"
            "// setup ASIO sound card for the subset 1\n"
            "/////////////////////////////////////////////////////\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device = Asio\n"
            "command = blocking\n"
            "value = 1\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device = Asio\n"
            "command = autodetect\n"
            "value = 0\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  Asio\n"
            "command = rec_ch_mask\n"
            "value = 0x" + mask + "\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description = Input play samplerate\n"
            "device =  Asio\n"
            "command = play_samplerate\n"
            "value = 192000\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description = Output record samplerate\n"
            "device =  Asio\n"
            "command = record_samplerate\n"
            )
            if fs == "4fs" and legacy == 1:
                file.write(
                "value = 192000\n"
                )
            elif fs == "2fs" and legacy == 1:
                file.write(
                    "value = 96000\n"
                )
            elif fs == "1fs" and legacy == 1:
                file.write(
                    "value = 48000\n"
                )
            elif fs == "1fs" and legacy == 0:
                file.write(
                    "value = 48000\n"
                )
            elif fs == "2fs" and legacy == 0:
                file.write(
                    "value = 48000\n"
                )
            elif fs == "4fs" and legacy == 0:
                file.write(
                    "value = 48000\n"
                )
            file.write(
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description = 4 compressed input line\n"
            "device = Asio\n"
            "command = n_play_i2s_lines\n"
            "value = 4\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device = Asio\n"
            "command = samplesize\n"
            "value = 24\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device = Asio\n"
            "command = justification\n"
            "value = 1\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  Asio\n"
            "command = n_repeat_play_file\n"
            "value = 1\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description = Dolby_Soundbar stream type\n"
            "device = Asio\n"
            "command = play_file_type\n"
            "value = raw\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device = Asio\n"
            "command = generate_zeros_ms\n"
            "value = 1000\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  Asio\n"
            "command = grab_file\n"
            "value = StringManager_3::value\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  Asio\n"
            "command = play_file\n"
            "value = StringManager_4::value\n"
            "delay = 0\n"
            "option = [SET]\n"
            "\n"
            )
            file.write(
            "////////////////////////////start playing/////////////////////////////////////////\n"
            "\n"
            "[step]\n"
            "description =\n"
            "device =  Asio\n"
            "command = run\n"
            "value = 1\n"
            "delay = 0\n"
            "option = [GET]\n"
            "\n"
            )
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


