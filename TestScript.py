import unittest
from unittest import mock
import os
import shutil
import tempfile
import Script_Copiado_Impresion3D

# class TestCopyFolder(unittest.TestCase):
#     def setUp(self):
#         self.source_dir = tempfile.mkdtemp()
#         self.destination_dir = tempfile.mkdtemp()
#         self.source_file = os.path.join(self.source_dir, "test.txt")
#         self.destination_file = os.path.join(self.destination_dir, "test.txt")
#         with open(self.source_file, "w") as f:
#             f.write("Test data")

#     def test_copy_folder(self):
#         Script_Copiado_Impresion3D.copy_folder(self.source_dir, self.destination_dir)
#         self.assertTrue(os.path.exists(self.destination_dir))
#         self.assertTrue(os.path.exists(self.destination_file))
#         with open(self.destination_file, "r") as f:
#             content = f.read()
#         self.assertEqual(content, "Test data")

#     def tearDown(self):
#         shutil.rmtree(self.source_dir)
#         shutil.rmtree(self.destination_dir)

class TestAddFormatedDataToCSV(unittest.TestCase):
    def test_addFormatedDataToCSV(self):
        source_item = "12345_file.gcode"
        nombrePadre = "C:\\what\\ProjectName\\12345_file.gcode"
        Script_Copiado_Impresion3D.addFormatedDataToCSV(source_item, nombrePadre)
        expected_data = [["ProjectName", "12345", "file.gcode"]]
        self.assertEqual(Script_Copiado_Impresion3D.datos_archivos, expected_data)

if __name__ == '__main__':
    unittest.main()