from PyQt5 import QtWidgets, QtCore, QtGui
import design
import win32clipboard
import os
import subprocess
import numpy as np
import cv2
import imutils
from PIL import Image
from skimage.filters import threshold_local
import fitz
import io
import ctypes
from data import image_path

def order_points(pts):
        rect = np.zeros((4, 2), dtype = "float32")
        s = pts.sum(axis = 1)
        rect[0] = pts[np.argmin(s)]
        rect[2] = pts[np.argmax(s)]
        diff = np.diff(pts, axis = 1)
        rect[1] = pts[np.argmin(diff)]
        rect[3] = pts[np.argmax(diff)]
        return rect


def four_point_transform(image, pts):
    rect = order_points(pts)
    (tl, tr, br, bl) = rect
    widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
    widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
    maxWidth = max(int(widthA), int(widthB))
    heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
    heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
    maxHeight = max(int(heightA), int(heightB))
    dst = np.array([[0, 0], [maxWidth - 1, 0], [maxWidth - 1, maxHeight - 1], [0, maxHeight - 1]], dtype = "float32")
    M = cv2.getPerspectiveTransform(rect, dst)
    warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))
    return warped

class Main_app(QtWidgets.QMainWindow, design.Ui_MainWindow):

    tray_icon = None                                                            #Initializing tray icon
    popup_is_shown = True
    is_from_tray = False

    def open_ui_tab(self, index):
        self.stacked_widget.setCurrentIndex(index)


    def resolve_tray_activation_reason(self, reason):                           #Resolves the way the tray was activated, opens the app if it was double-clicked
        if reason == QtWidgets.QSystemTrayIcon.DoubleClick:
            self.show_from_tray()


    def close_from_tray(self):                                                  #Check whether the command to close the app is from the tray or from the app
        self.is_from_tray = True
        self.close()


    def show_from_tray(self):
        self.tray_icon.hide()
        self.show()


    def stop_showing_popup(self):
        self.popup_is_shown = False


    def closeEvent(self, event):                                                #Redefining closeEvent to change behaviour depending on closure caller
        if self.is_from_tray:
            event.accept()
        else:
            event.ignore()
            self.tray_icon.show()
            if self.popup_is_shown:
                self.tray_icon.showMessage('Уведомление', 'Приложение было скрыто в трей. Чтобы уведомление более не отображалось, нажмите на него.',
                                            QtGui.QIcon(image_path), 10000) #Shows a message informing about the app being closed (hidden)
            self.hide()

    def picture_to_scan(self):
        for i in range(self.photo_to_scan_list_main.count()):
            item = self.photo_to_scan_list_main.item(i).text()
            existence = os.path.exists(item)
            if not existence:
                continue

            photo_path = item
            if not self.photo_to_scan_end_name.text():
                photo_dir = os.path.dirname(photo_path)
            else:
                photo_dir = os.path.dirname(self.photo_to_scan_end_name.text())

            result_name = os.path.basename(item)[:-4] + ".png"

            image = cv2.imread(photo_path)
            ratio = image.shape[0]/500.0
            original = image.copy()
            image = imutils.resize(image, height = 500)

            grayscale = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            grayscale = cv2.GaussianBlur(grayscale, (5, 5), 0)
            edged = cv2.Canny(grayscale, 75, 200)


            #print('Edge detection:')
            #cv2.imshow('Image', image)
            #cv2.imshow('Edged', edged)
            #cv2.waitKey(0)
            #cv2.destroyAllWindows()

            contours = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
            contours = imutils.grab_contours(contours)
            contours = sorted(contours, key = cv2.contourArea, reverse = True)[:5]

            for contour in contours:
                perimeter = cv2.arcLength(contour, True)
                approx = cv2.approxPolyDP(contour, 0.02 * perimeter, True)

                if len(approx) == 4:
                    screen_contour = approx
                    break

            #print('Finding contours:')
            #cv2.drawContours(image, [screen_contour], -1, (0, 255, 0), 2)
            #cv2.imshow('Outline', image)
            #cv2.waitKey(0)
            #cv2.destroyAllWindows()

            try:
                warped = four_point_transform(original, screen_contour.reshape(4, 2) * ratio)
            except NameError:
                continue
            else:
                warped = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
                T = threshold_local(warped, 11, offset = 10, method = "gaussian")
                warped = (warped > T).astype("uint8") * 255

            #print('Getting scanned picture:')
            #cv2.imshow('Original:', imutils.resize(original, height = 650))
            #cv2.imshow('Scanned:', imutils.resize(warped, height = 650))
            #cv2.waitKey(0)
            #cv2.destroyAllWindows()
                #imwrite is still a part of 'else' statement!
                cv2.imwrite(photo_dir + '\\' + result_name, imutils.resize(warped, height = 650))


    def call_diffuse(self):
        path_to_first = self.doc_comp_file_1.text()
        path_to_second = self.doc_comp_file_2.text()
        os.system('diffuse "' + path_to_first + '" "' + path_to_second + '"')


    def get_clipboard_data(self):
        win32clipboard.OpenClipboard()
        try:
            filenames = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
        except TypeError:
            win32clipboard.CloseClipboard()
            return []
        else:
            win32clipboard.CloseClipboard()
            return filenames


    def write_path_to_line(self, data, item, is_file):
        if is_file and os.path.isfile(data):
            item.setText(data)
        elif not (is_file) and os.path.isdir(data):
            item.setText(data)


    def write_items_to_list(self, data, item, extension_allowed): #0 - all extensions, 1 - djvu, 2 - pictures(png/jpg), 3 - pdf
        if extension_allowed == 0:
            for i in data:
                item.addItem(i)

        elif extension_allowed == 1:
            for i in data:
                if i[-5:] == '.djvu':
                    item.addItem(i)

        elif extension_allowed == 2:
            for i in data:
                checked_extension = i[-4:]  #We don't want Python to create and then delete two separate instances if i[-4:]
                if checked_extension == '.png' or checked_extension == '.jpg':
                    item.addItem(i)

        elif extension_allowed == 3:
            for i in data:
                if i[-4:] == '.pdf':
                    item.addItem(i)


    def shortcut_process(self, number): #number defines, what shortcut it was: 0 - first one, 1 - second one
        page_index = self.stacked_widget.currentIndex()
        clipboard_data = self.get_clipboard_data()
        data_length = len(clipboard_data)
        if data_length != 0:    #If there are no items in the clipboard, we don't need to do anything

            if page_index == 0:
                if data_length == 2:
                    self.write_path_to_line(clipboard_data[0], self.doc_comp_file_1, True)
                    self.write_path_to_line(clipboard_data[1], self.doc_comp_file_2, True)
                elif number == 0 and data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.doc_comp_file_1, True)
                elif data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.doc_comp_file_2, True)

            elif page_index == 1:
                if number == 0 and data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.repo_get_repo_folder_name, False)
                elif data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.repo_get_end_folder_name, False)

            elif page_index == 2:
                if number == 0 and data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.git_commit_repo_name, False)
                else:
                    self.write_items_to_list(clipboard_data, self.git_commit_list_main, 0)

            elif page_index == 3:
                if number == 0:
                    self.write_items_to_list(clipboard_data, self.convert_djvu_list_main, 1)
                elif data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.convert_djvu_end_name, False)

            elif page_index == 4:
                if number == 0:
                    self.write_items_to_list(clipboard_data, self.photo_to_scan_list_main, 2)
                elif data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.photo_to_scan_end_name, False)

            elif page_index == 5:
                if number == 0:
                    self.write_items_to_list(clipboard_data, self.pdf_from_pic_list_main, 2)
                elif data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.pdf_from_pic_end_name, False)

            elif page_index == 6:
                if number == 0:
                    self.write_items_to_list(clipboard_data, self.extract_from_pics_list_main, 3)
                elif number == 1 and data_length == 1:
                    self.write_path_to_line(clipboard_data[0], self.extract_from_pics_end_name, False)

            elif page_index == 7:
                self.write_items_to_list(clipboard_data, self.pdf_ocr_list_main, 3)


    def button_add_function(self, item, is_line, is_file = True, extension_allowed = 0):
        if is_file:
            data = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите файл...")[0]
        else:
            data = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите папку...")

        if is_line and data:
            self.write_path_to_line(data, item, is_file)
        elif data:
            self.write_items_to_list([data], item, extension_allowed)


    def eventFilter(self, source, event):
        if event.type() == QtCore.QEvent.MouseMove and event.buttons() == QtCore.Qt.LeftButton:
            self.move(event.globalPos().x() - self.old_pos_x, event.globalPos().y() - self.old_pos_y)

        elif event.type() == QtCore.QEvent.MouseButtonPress and event.buttons() == QtCore.Qt.LeftButton:
            self.old_pos_x = event.localPos().x()
            self.old_pos_y = event.localPos().y()
        return True


    def move_hash(self):
        if self.repo_get_ver_list_main.currentIndex() != -1:
            shortened_hash = self.repo_get_ver_list_main.currentItem().text().split()[0]
            self.repo_get_ver_list_name.setText(shortened_hash)


    def get_version_list(self):
        self.repo_get_ver_list_main.clear()
        a = 'cd /d ' + self.repo_get_repo_folder_name.text() + ' '
        b = ' git log --pretty=format:"%h - %as %s" > tmp_file.txt'
        os.system(a + '&' + b)                                              #Acquiring all the logs in format "shortened_hash - commit_date commit_msg" and put it in tmp_file.txt
        file_name = self.repo_get_repo_folder_name.text() + '\\tmp_file.txt'
        tmp_file = open(file_name, 'r', encoding="utf-8")
        version_list = [line.rstrip() for line in tmp_file]                 #version_list is a list of git logs, we need rstrip() because every line has /n at the end otherwise
        tmp_file.close()
        c = 'del tmp_file.txt'
        os.system(a + '&' + c)                                              #Removing the temporary file
        for line in version_list:
            self.repo_get_ver_list_main.addItem(line)


    def mirror_commit(self):
        start_path = self.repo_get_repo_folder_name.text()
        version = self.repo_get_ver_list_name.text()
        end_path = self.repo_get_end_folder_name.text()
        onlyfiles = [f for f in os.listdir(start_path) if os.path.isfile(os.path.join(start_path, f))]

        a = 'cd /d "' + start_path + '" '
        b = ' git checkout ' + version + ' '
        c = ' copy ' + '\"'
        d = '\" \"' + self.control_end_name.text() + '\" '
        e = ' git checkout master'

        os.system(a + "&" + b)
        for item in onlyfiles:
            os.system(c + item + d)
        os.system(a + "&" + e)


    def create_commit(self):
        repo_path = self.git_commit_repo_name.text()
        onlyfiles = [f for f in os.listdir(repo_path) if os.path.isfile(repo_path)]
        commit_msg = self.pdf_from_pic_commit_name.text()

        a = 'cd /d "' + repo_path + '" '
        b = ' git init '
        c = ' copy \"'
        d = ' replace \"' + source_file + '\" '
        e = ' git add . '
        f = ' git commit -m "' + commit_msg + '"'

        sp_array = ['git', '-C', repo_path, 'rev-parse']
        is_new = subprocess.run(sp_array, stdout = subprocess.DEVNULL, stderr = subprocess.DEVNULL, encoding = 'utf-8').returncode
        if is_new:
            os.system(a + '&' + b)

        for i in range(self.git_commit_list_main.count()):
            item = self.git_commit_list_main.item(i).text()
            if os.path.basename(item) in os.path.basename(onlyfiles):
                os.system(a + "&" + d + item + '\" ')
            else:
                os.system(a + "&" + c + item + '\" ')
        os.system(a + '&' + e + '&' + f)


    def convert_djvu_to_pdf(self):
        if self.convert_djvu_end_name.text():
            result_folder = os.dirname(self.convert_djvu_end_name.text()) + '\\'
            for i in range(self.convert_djvu_list_main.count()):
                item = self.convert_djvu_list_main.item(i).text()
                result_path = result_folder + os.path.basename(item)[:-5] + '.pdf'
                os.system('ddjvu -format=pdf "' + item + '" "' + result_path + '"')
        else:
            for i in range(self.convert_djvu_list_main.count()):
                item = self.convert_djvu_list_main.item(i).text()
                result_folder = os.dirname(item) + '\\'
                result_path = result_folder + os.path.basename(item)[:-5] + '.pdf'
                os.system('ddjvu -format=pdf "' + item + '" "' + result_path + '"')


    def compile_pdf_from_photos(self):
        if self.pdf_from_pic_list_main.count() >= 1:
            photo_path = self.pdf_from_pic_list_main.item(0).text()
            if self.pdf_from_pic_end_name.text():
                photo_dir = self.pdf_from_pic_end_name.text()
            else:
                photo_dir = os.path.dirname(photo_path)
            if not self.pdf_from_pic_endname_name.text():
                self.pdf_from_pic_endname_name.setText('default_converted_name')
            result_name = self.pdf_from_pic_endname_name.text() + ".pdf"
            first_picture = Image.open(self.pdf_from_pic_list_main.item(0).text()).convert('RGB')
            if self.pdf_from_pic_list_main.count() == 1:
                first_picture.save(photo_dir + '\\' + result_name)
            else:
                length_of_list = self.pdf_from_pic_list_main.count()
                added_photos = [Image.open(self.pdf_from_pic_list_main.item(i).text()).convert('RGB') for i in range(1, length_of_list)]
                print (photo_dir + '\\' + result_name)
                print ('path - dirname: ' + os.path.dirname(photo_path))
                first_picture.save(photo_dir + '\\' + result_name, save_all = True, append_images = added_photos)


    def extract_pictures_from_pdf(self):
        for i in range(self.extract_from_pics_list_main.count()):
            file = self.extract_from_pics_list_main.item(i).text()
            file_short = os.path.basename(file)[:-4]
            if self.extract_from_pics_end_name.text():
                end_path = self.extract_from_pics_end_name.text()
            else:
                end_path = os.dirname(file)
            pdf_file = fitz.open(file)
            for page_index in range(len(pdf_file)):
                page = pdf_file[page_index]
                image_list = page.getImageList()
                if not image_list:
                    continue
                for image_index, img in enumerate(page.getImageList(), start=1):
                    xref = img[0]
                    base_image = pdf_file.extractImage(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    image = Image.open(io.BytesIO(image_bytes))
                    image.save(open(end_path + '\\' + file_short + str(page_index+1) + '_' + str(image_index) + '.' + image_ext, "wb"))


    def ocr_pdf(self):
        for i in range(self.pdf_ocr_list_main.count()):
            item = self.pdf_ocr_list_main.item(i).text()
            item_renamed = item[:-4] + '_renamed.pdf'
            if not item:
                continue
            ru_added = self.pdf_ocr_setup_lang_ru.isChecked()
            en_added = self.pdf_ocr_setup_lang_en.isChecked()
            added = ''
            if ru_added:
                added = '-l rus'
                if en_added:
                    added += '+eng'
            else:
                if en_added:
                    added = '-l eng'
            is_force = self.pdf_ocr_setup_ignore_check.isChecked()
            if is_force:
                added += ' -f'
            is_txt = self.pdf_ocr_setup_txt.isChecked()
            added += ' '
            os.system('ocrmypdf ' + added + ' "' + item + '" "' + item_renamed + '"')
            if is_txt:
                added = added.replace('-f', '')
                item_renamed = item_renamed[:-4] + '.txt'
                os.system('tesseract ' + added + ' "' + item + '" "' + item_renamed + '"')


    def get_startup_path(self):                                                     #Returns the path to the startup user folder (we can't be sure that main drive is C:)
        return '"' + os.getenv('SystemDrive') + os.getenv('HOMEPATH') + r'\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Text_Multitool.exe"'


    def add_to_startup(self):                                                       #Adds the application to the startup by placing the shortcut in the startup folder
        inp_str = 'mklink ' + self.get_startup_path() + ' "' + sys.executable + '"'
        os.system(inp_str)


    def delete_from_startup(self):                                                  #Removes the application from the startup by removing the shortcut from the startup folder (if it exists)
        inp_str = 'del ' + self.get_startup_path()
        os.system(inp_str)


    def __init__(self):
        super().__init__()
        self.setupUi(self)
        #Left side buttons
        self.compare_docs                 .clicked.connect(lambda: self.open_ui_tab(0))
        self.git_get_version              .clicked.connect(lambda: self.open_ui_tab(1))
        self.create_git_ver               .clicked.connect(lambda: self.open_ui_tab(2))
        self.convert_djvu                 .clicked.connect(lambda: self.open_ui_tab(3))
        self.convert_photo                .clicked.connect(lambda: self.open_ui_tab(4))
        self.create_pdf_from_pic          .clicked.connect(lambda: self.open_ui_tab(5))
        self.extract_pics                 .clicked.connect(lambda: self.open_ui_tab(6))
        self.recognize_text               .clicked.connect(lambda: self.open_ui_tab(7))
        self.settings_about               .clicked.connect(lambda: self.open_ui_tab(8))

        #Main frame buttons
        #First Tab
        self.doc_comp_button_1            .clicked.connect(lambda: self.button_add_function(self.doc_comp_file_1, True))
        self.doc_comp_button_2            .clicked.connect(lambda: self.button_add_function(self.doc_comp_file_2, True))
        self.doc_comp_do                  .clicked.connect(self.call_diffuse)
        #Second Tab
        self.repo_get_repo_folder_button  .clicked.connect(lambda: self.button_add_function(self.repo_get_repo_folder_name, True, False))
        self.repo_get_end_folder_button   .clicked.connect(lambda: self.button_add_function(self.repo_get_end_folder_name, False))
        self.repo_get_ver_list_button     .clicked.connect(self.move_hash)
        self.repo_get_ver_list_load       .clicked.connect(self.get_version_list)
        self.repo_get_do_button           .clicked.connect(self.mirror_commit)
        #Third Tab
        self.git_commit_repo_button       .clicked.connect(lambda: self.button_add_function(self.git_commit_repo_name, True, False))
        self.git_commit_list_add          .clicked.connect(lambda: self.button_add_function(self.git_commit_list_main, False))
        self.git_commit_list_delete       .clicked.connect(lambda: self.git_commit_list_main.takeItem(self.git_commit_list_main.currentRow()))
        self.git_commit_list_clear        .clicked.connect(self.git_commit_list_main.clear)
        self.git_commit_do_button         .clicked.connect(self.create_commit)
        #Fourth tab
        self.convert_djvu_list_add        .clicked.connect(lambda: self.button_add_function(self.convert_djvu_list_main, False))
        self.convert_djvu_list_delete     .clicked.connect(lambda: self.convert_djvu_list_main.takeItem(self.convert_djvu_list_main.currentRow()))
        self.convert_djvu_list_clear      .clicked.connect(self.convert_djvu_list_main.clear)
        self.convert_djvu_end_button      .clicked.connect(lambda: self.button_add_function(self.convert_djvu_end_name, False, True, 1))
        self.convert_djvu_do_button       .clicked.connect(self.convert_djvu_to_pdf)
        #Fifth Tab
        self.photo_to_scan_list_add       .clicked.connect(lambda: self.button_add_function(self.photo_to_scan_list_main, False, True, 2))
        self.photo_to_scan_list_delete    .clicked.connect(lambda: self.photo_to_scan_list_main.takeItem(self.photo_to_scan_list_main.currentRow()))
        self.photo_to_scan_list_clear     .clicked.connect(self.photo_to_scan_list_main.clear)
        self.photo_to_scan_end_button     .clicked.connect(lambda: self.button_add_function(self.photo_to_scan_end_name, True, False))
        self.photo_to_scan_do_button      .clicked.connect(self.picture_to_scan)
        #Sixth Tab
        self.pdf_from_pic_list_add        .clicked.connect(lambda: self.button_add_function(self.pdf_from_pic_list_main, False, True, 2))
        self.pdf_from_pic_list_delete     .clicked.connect(lambda: self.pdf_from_pic_list_main.takeItem(self.pdf_from_pic_list_main.currentRow()))
        self.pdf_from_pic_list_clear      .clicked.connect(self.pdf_from_pic_list_main.clear)
        self.pdf_from_pic_end_button      .clicked.connect(lambda: self.button_add_function(self.pdf_from_pic_end_name, True, False))
        self.pdf_from_pic_do_button       .clicked.connect(self.compile_pdf_from_photos)
        #Seventh Tab
        self.extract_from_pics_list_add   .clicked.connect(lambda: self.button_add_function(self.extract_from_pics_list_main, False, True, 3))
        self.extract_from_pics_list_delete.clicked.connect(lambda: self.extract_from_pics_list_main.takeItem(self.extract_from_pics_list_main.currentRow()))
        self.extract_from_pics_list_clear .clicked.connect(self.extract_from_pics_list_main.clear)
        self.extract_from_pics_end_button .clicked.connect(lambda: self.button_add_function(self.extract_from_pics_end_name, True, False))
        self.extract_from_pics_do_button  .clicked.connect(self.extract_pictures_from_pdf)
        #Eighth Tab
        self.pdf_ocr_list_add             .clicked.connect(lambda: self.button_add_function(self.pdf_ocr_list_main, False, True, 3))
        self.pdf_ocr_list_delete          .clicked.connect(lambda: self.pdf_ocr_list_main.takeItem(self.pdf_ocr_list_main.currentRow()))
        self.pdf_ocr_list_clear           .clicked.connect(self.pdf_ocr_list_main.clear)
        self.pdf_ocr_do_button            .clicked.connect(self.ocr_pdf)
        #Settings Tab
        self.settings_autoboot_add        .clicked.connect(self.add_to_startup)
        self.settings_autoboot_delete     .clicked.connect(self.delete_from_startup)

        #Shortcut processing
        QtWidgets.QShortcut(QtGui.QKeySequence('Ctrl+1'),self).activated.connect(lambda: self.shortcut_process(0))
        QtWidgets.QShortcut(QtGui.QKeySequence('Ctrl+2'),self).activated.connect(lambda: self.shortcut_process(1))

        #Close and minimize buttons
        self.close_button.clicked.connect(self.close)
        self.minimize_button.clicked.connect(self.showMinimized)

        #Event filtering of the top frame to allow window dragging
        self.top_frame.installEventFilter(self)

        #Tray menu and icon: setting up
        if os.path.exists(image_path):
            self.tray_icon = QtWidgets.QSystemTrayIcon(self)
            self.tray_icon.setIcon(QtGui.QIcon(image_path))                     #Picture has to be in folder of the program
        #Creating actions
            self.show_action = QtWidgets.QAction("Show", self)
            self.quit_action = QtWidgets.QAction("Exit", self)
        #Connecting actions to functions
            self.show_action.triggered.connect(self.show_from_tray)                 #Action to show the main program window
            self.quit_action.triggered.connect(self.close_from_tray)                #Action to exit the program. Sets is_from_tray parameter to True, needed to distinguish closeEvent source
        #Functions and setting up menu
            self.tray_menu = QtWidgets.QMenu()
            self.tray_menu.addAction(self.show_action)
            self.tray_menu.addAction(self.quit_action)
            self.tray_icon.setContextMenu(self.tray_menu)
        #Show the app window on tray icon double click
            self.tray_icon.activated.connect(self.resolve_tray_activation_reason)
        #Stop showing popup notification on app hiding if one is clicked
            self.tray_icon.messageClicked.connect(self.stop_showing_popup)


def main():
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(' ')
    app = QtWidgets.QApplication([])
    window = Main_app()
    window.setWindowFlags(QtCore.Qt.FramelessWindowHint)
    window.show()
    app.setWindowIcon(QtGui.QIcon(image_path))
    app.exec_()

if __name__ == '__main__':
    main()
