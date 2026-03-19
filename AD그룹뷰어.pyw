# exe 패키징 시 주의
# ldap3의 일부 모듈(ldap3.utils.conv, ldap3.core.exceptions)은 자동 포함되지 않음
# → auto-py-to-exe에서 hidden import에 반드시 추가해야 함

import sys
import tempfile
import re

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QLineEdit, QAbstractItemView,
    QTextEdit, QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QFormLayout, QGridLayout,
    QMessageBox, QComboBox, QProgressDialog, QHBoxLayout, QFileDialog, QInputDialog, QCompleter,
    QListWidget, QListWidgetItem, QScrollArea
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5 import QtCore
from ldap3.utils.conv import escape_filter_chars
from ldap3.core.exceptions import LDAPInvalidFilterError
import ldap3
import json
import os
import subprocess
from dataclasses import dataclass


@dataclass
class MemberInfo:
    sAMAccountName: str
    department: str
    displayName: str
    mail: str

class ADGroupViewer(QWidget):
    def __init__(self):
        super().__init__()
        self.account_info = self.load_account_info()
        if not isinstance(self.account_info, dict):
            self.account_info = {"server_ip": "", "user": "", "password": ""}
        self.member_list = []
        self.initUI()
        self.apply_custom_styles()
 
    def apply_custom_styles(self):
        custom_style = """
            QWidget {
                background-color: #2b2b2b;
                color: #e0e0e0;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
            }

            QPushButton {
                background-color: #3C3C3C;
                color: #E0E0E0;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #2A2A2A;
            }

            QLabel {
                font-size: 13px;
                font-weight: bold;
                color: #f0f0f0;
            }

            QLineEdit {
                background-color: #3b3b3b;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 5px;
                color: #f0f0f0;
            }

            QTableWidget {
                background-color: #3b3b3b;
                border: 1px solid #444;
                gridline-color: #555;
            }

            QHeaderView::section {
                background-color: #3C3C3C;
                color: white;
                padding: 8px;
                border: 1px solid #2A2A2A;
            }

            QTextEdit {
                background-color: #3b3b3b;
                border: 1px solid #555;
                padding: 5px;
                color: #f0f0f0;
            }
        """
        self.setStyleSheet(custom_style)

    def style_buttons(self, button):
        button.setFixedSize(110, 40)
        button.setStyleSheet("""
            QPushButton {
                background-color: #3C3C3C;
                color: #E0E0E0;
                border-radius: 6px;
                padding: 10px;
                text-align: center;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2A2A2A;
            }
        """)

    def initUI(self):
        self.setWindowTitle('AD 그룹 뷰어')
        self.setGeometry(100, 100, 650, 700)
        self.layout = QVBoxLayout()

        search_layout = QGridLayout()
        self.group_label = QLabel('AD 그룹 이름:')
        self.group_input = QLineEdit()

        self.show_button = QPushButton('검색')
        self.create_group_button = QPushButton('그룹 생성')
        self.style_buttons(self.show_button)
        self.style_buttons(self.create_group_button)

        search_layout.addWidget(self.group_label, 0, 0)
        search_layout.addWidget(self.group_input, 0, 1, 1, 2)
        search_layout.addWidget(self.show_button, 0, 3)
        search_layout.addWidget(self.create_group_button, 0, 4)

        filter_layout = QGridLayout()
        self.filter_label = QLabel("결과내 검색:")
        self.filter_input = QLineEdit()
        self.copy_button = QPushButton('결과 복사')
        self.save_button = QPushButton('내보내기')

        self.style_buttons(self.copy_button)
        self.style_buttons(self.save_button)

        filter_layout.addWidget(self.filter_label, 0, 0)
        filter_layout.addWidget(self.filter_input, 0, 1, 1, 2)
        filter_layout.addWidget(self.copy_button, 0, 3)
        filter_layout.addWidget(self.save_button, 0, 4)

        self.filter_input.textChanged.connect(self.filter_member_table)

        management_layout = QHBoxLayout()
        self.account_button = QPushButton('AD 서버 정보')
        self.group_manage_button = QPushButton('그룹 관리')
        self.manage_members_button = QPushButton('멤버 관리')
        self.sync_button = QPushButton('동기화')

        self.account_button.setFixedSize(150, 45)
        self.group_manage_button.setFixedSize(150, 45)
        self.manage_members_button.setFixedSize(150, 45)
        self.sync_button.setFixedSize(150, 45)

        management_layout.addStretch()
        management_layout.addWidget(self.account_button)
        management_layout.addSpacing(20)
        management_layout.addWidget(self.group_manage_button)
        management_layout.addSpacing(20)
        management_layout.addWidget(self.manage_members_button)
        management_layout.addSpacing(20)
        management_layout.addWidget(self.sync_button)
        management_layout.addStretch()

        self.member_table = QTableWidget()
        self.member_table.setMinimumHeight(300)
        self.member_table.setColumnCount(4)
        self.member_table.setHorizontalHeaderLabels(['사원 번호', '부서', '표시 이름', '메일 주소'])
        self.member_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.member_table.setRowCount(0)

        self.member_table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.member_table.verticalHeader().setDefaultSectionSize(40)
        self.member_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.member_table.verticalHeader().setVisible(False)
        self.member_table.setSortingEnabled(True)

        self.result_text = QTextEdit()
        self.result_text.setFixedHeight(150)

        self.layout.addLayout(search_layout)
        self.layout.addLayout(filter_layout)
        self.layout.addWidget(self.member_table)
        self.layout.addLayout(management_layout)
        self.layout.addWidget(self.result_text)

        self.setLayout(self.layout)
        self.center()

        self.show_button.clicked.connect(self.show_group_members)
        self.create_group_button.clicked.connect(self.open_group_creation_dialog)
        self.copy_button.clicked.connect(self.copy_all_members)
        self.save_button.clicked.connect(self.save_member_list)
        self.account_button.clicked.connect(self.open_account_management)
        self.group_manage_button.clicked.connect(self.open_group_management)
        self.manage_members_button.clicked.connect(self.open_member_management)
        self.sync_button.clicked.connect(self.execute_sync_script)

    def make_center_item(self, value):
        item = QTableWidgetItem(str(value) if value is not None else "")
        item.setTextAlignment(Qt.AlignCenter)
        return item

    def center_on_parent(self, dialog):
        parent_geometry = self.geometry()
        dialog_geometry = dialog.geometry()

        x = parent_geometry.center().x() - (dialog_geometry.width() // 2)
        y = parent_geometry.center().y() - (dialog_geometry.height() // 1)

        dialog.move(x, y)

    def open_group_creation_dialog(self):
        group_name = self.group_input.text().strip()
        server_ip = self.account_info.get('server_ip', '').strip()
        user = self.account_info.get('user', '').strip()
        password = self.account_info.get('password', '').strip()

        if not server_ip or not user or not password:
            QMessageBox.critical(self, "오류", "AD 서버 정보가 입력되지 않았습니다.\n관리자 계정 관리에서 설정해 주세요.")
            return

        server_uri = f"ldap://{server_ip}"
        try:
            conn = ldap3.Connection(server_uri, user=user, password=password, auto_bind=True)
            conn.unbind()
        except ldap3.core.exceptions.LDAPBindError:
            QMessageBox.critical(self, "AD 인증 실패", "AD 서버 인증에 실패했습니다.\nID 또는 비밀번호를 확인해 주세요.")
            return
        except ldap3.LDAPException as e:
            QMessageBox.critical(self, "AD 연결 오류", f"AD 서버 연결 중 오류 발생:\n{str(e)}")
            return

        if not group_name:
            QMessageBox.warning(self, "경고", "그룹 이름을 입력하세요.")
            return
        
        try:
            dialog = CreateGroupDialog(group_name, self.account_info, self)
            self.center_on_parent(dialog)
            if dialog.exec_():
                QMessageBox.information(self, "성공", f"그룹 '{group_name}'이 생성되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"그룹 생성 중 오류 발생:\n{str(e)}")

    def save_member_list(self):
        if not self.member_list:
            QMessageBox.warning(self, "경고", "저장할 멤버 목록이 없습니다.")
            return

        options = QFileDialog.Options() | QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "멤버 목록 저장",
            os.path.join(os.getcwd(), "member_list.csv"),
            "CSV Files (*.csv);;All Files (*)",
            options=options
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8-sig') as file:
                    file.write("사원 번호,부서,표시 이름,메일 주소\n")
                    for member in self.member_list:
                        sAMAccountName = member.sAMAccountName if member.sAMAccountName else ""
                        department = member.department if member.department else ""
                        displayName = member.displayName if member.displayName else ""
                        mail = member.mail if member.mail else ""
                        line = f"{sAMAccountName},{department},{displayName},{mail}\n"
                        file.write(line)

                QMessageBox.information(self, "성공", "멤버 목록이 성공적으로 저장되었습니다.")

            except Exception as e:
                QMessageBox.critical(self, "오류", f"파일 저장 중 오류 발생:\n{str(e)}")

    def filter_member_table(self):
        filter_text = self.filter_input.text().strip().lower()
        for row in range(self.member_table.rowCount()):
            match_found = False
            for col in range(self.member_table.columnCount()):
                item = self.member_table.item(row, col)
                if item and filter_text in item.text().strip().lower():
                    match_found = True
                    break
            self.member_table.setRowHidden(row, not match_found)

    def center(self):
        screen = QApplication.primaryScreen()
        if screen:
            screen_geometry = screen.availableGeometry()
            self.move(screen_geometry.center() - self.rect().center())

    def execute_sync_script(self):
        user = self.account_info.get('user')
        password = self.account_info.get('password')

        confirm = QMessageBox.question(
            self, 
            "동기화 확인", 
            "정말로 동기화를 실행하시겠습니까?", 
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )

        if confirm != QMessageBox.Yes:
            return

        if not user or not password:
            QMessageBox.warning(self, "경고", "관리자 계정 정보가 없습니다. 관리자 계정 관리를 통해 정보를 설정해주세요.")
            return

        try:
            self.run_powershell_sync(user, password)
            QMessageBox.information(self, "성공", "동기화가 성공적으로 완료되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"동기화 중 오류가 발생했습니다:\n{str(e)}")

    def run_powershell_sync(self, user, password):
        computer_name = "Msync.lskglobal.com"

        ps_command = f'''
        $User = "{user}"
        $PWord = ConvertTo-SecureString -String "{password}" -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord
        Invoke-Command -ComputerName "{computer_name}" -Credential $Credential -ScriptBlock {{
            Start-ADSyncSyncCycle -PolicyType Delta
        }}
        '''

        try:
            subprocess.run(
                ["powershell.exe", "-Command", ps_command],
                check=True,
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

        except subprocess.CalledProcessError as e:
            error_output = e.stderr or str(e)
            QMessageBox.critical(self, "동기화 실패", f"PowerShell 실행 중 오류 발생:\n{error_output}")

    def load_account_info(self):
        path = "C:/account_info.json"
        try:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
        except Exception:
            pass
        return {"server_ip": "", "user": "", "password": ""}

    def save_account_info(self, server_ip, user, password):
        self.account_info = {'server_ip': server_ip, 'user': user, 'password': password}
        path = "C:/account_info.json"
        with open(path, 'w') as f:
            json.dump(self.account_info, f)

    def open_account_management(self):
        dialog = AccountManagementDialog(
            self.account_info.get('server_ip', ''),
            self.account_info.get('user', ''),
            self.account_info.get('password', '')
        )
        self.center_on_parent(dialog)
        if dialog.exec_():
            new_server_ip, new_user, new_password = dialog.get_account_info()
            self.save_account_info(new_server_ip, new_user, new_password)

    def open_group_management(self):
        try:
            group_name = self.group_input.text().strip()
            server_ip = self.account_info.get('server_ip', '').strip()
            user = self.account_info.get('user', '').strip()
            password = self.account_info.get('password', '').strip()

            if not server_ip or not user or not password:
                QMessageBox.critical(self, "오류", "AD 서버 정보가 입력되지 않았습니다.\n관리자 계정 관리에서 설정해 주세요.")
                return

            if not group_name:
                QMessageBox.warning(self, "경고", "그룹 이름을 먼저 검색해 주세요.")
                return

            dialog = GroupManagementDialog(group_name, self.account_info, self)
            self.center_on_parent(dialog)
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"그룹 관리 열기 중 오류 발생:\n{str(e)}")

    def open_member_management(self):
        try:
            group_name = self.group_input.text().strip()
            server_ip = self.account_info.get('server_ip', '').strip()
            user = self.account_info.get('user', '').strip()
            password = self.account_info.get('password', '').strip()

            if not server_ip or not user or not password:
                QMessageBox.critical(self, "오류", "AD 서버 정보가 입력되지 않았습니다.\n관리자 계정 관리에서 설정해 주세요.")
                return

            if not group_name:
                 QMessageBox.warning(self, "경고", "그룹 이름을 입력해 주세요.")
                 return

            success = self.show_group_members()
            if not success:
                return

            if "그룹을 찾을 수 없습니다." in self.result_text.toPlainText():
                QMessageBox.warning(self, "경고", f"그룹 '{group_name}'을(를) 찾을 수 없습니다.")
                return

            dialog = MemberManagementDialog(group_name, self.member_list, self.account_info)
            self.center_on_parent(dialog)
            dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "오류", f"멤버 관리 열기 중 오류 발생:\n{str(e)}")

    def show_group_members(self):
        server_ip = self.account_info.get('server_ip', '').strip()
        user = self.account_info.get('user', '').strip()
        password = self.account_info.get('password', '').strip()
        group_name = self.group_input.text().strip()

        if not server_ip or not user or not password:
            QMessageBox.critical(self, "오류", "AD 서버 정보가 입력되지 않았습니다.\n관리자 계정 관리에서 설정해 주세요.")
            return

        if not group_name:
            QMessageBox.warning(self, "경고", "그룹 이름 또는 표시 이름을 입력하세요.")
            return

        server_uri = f"ldap://{server_ip}"
        self.member_list = []
        self.member_table.clearSpans()
        self.member_table.setSortingEnabled(False)
        self.member_table.setRowCount(0)

        try:
            with ldap3.Connection(server_uri, user=self.account_info['user'],
                                  password=self.account_info['password'], auto_bind=True) as conn:
                escaped_group_name = escape_filter_chars(group_name)
                search_filter = f"(|(cn={escaped_group_name})(displayName={escaped_group_name}))"
                conn.search(search_base="DC=lskglobal,DC=com", search_filter=search_filter, attributes=["member", "cn", "description"])
                entries = conn.entries

                if entries:
                    group = entries[0]
                    group_cn = str(group.cn) if hasattr(group, 'cn') else ""
                    group_description = group.description.value if hasattr(group, 'description') else ""
                    self.result_text.setPlainText(f"그룹 이름: {group_cn}\n그룹 설명: {group_description}\n\n")

                    members = [str(member) for member in group['member']]
                    if not members:
                        self.member_table.setRowCount(1)
                        self.member_table.setSpan(0, 0, 1, 4)
                        no_member_msg = QTableWidgetItem("그룹에 추가된 사용자가 없습니다.")
                        no_member_msg.setTextAlignment(Qt.AlignCenter)
                        self.member_table.setItem(0, 0, no_member_msg)
                        return True

                    member_attributes = {}
                    chunk_size = 50
                    for start in range(0, len(members), chunk_size):
                        chunk = members[start:start + chunk_size]
                        dn_filter = "".join(
                            f"(distinguishedName={escape_filter_chars(member_dn)})"
                            for member_dn in chunk
                        )
                        conn.search(
                            search_base="DC=lskglobal,DC=com",
                            search_filter=f"(|{dn_filter})",
                            attributes=["distinguishedName", "sAMAccountName", "department", "displayName", "mail"]
                        )
                        for user_entry in conn.entries:
                            member_attributes[user_entry.entry_dn] = user_entry

                    self.member_table.setRowCount(len(members))
                    for i, member_dn in enumerate(members):
                        user_attributes = member_attributes.get(member_dn)
                        if user_attributes:
                            sAMAccountName = user_attributes["sAMAccountName"][0] if "sAMAccountName" in user_attributes else ""
                            department = user_attributes["department"][0] if "department" in user_attributes else ""
                            displayName = user_attributes["displayName"][0] if "displayName" in user_attributes else ""
                            mail = user_attributes["mail"][0] if "mail" in user_attributes else ""
                            self.member_table.setItem(i, 0, self.make_center_item(sAMAccountName))
                            self.member_table.setItem(i, 1, self.make_center_item(department))
                            self.member_table.setItem(i, 2, self.make_center_item(displayName))
                            self.member_table.setItem(i, 3, self.make_center_item(mail))
                            self.member_list.append(MemberInfo(sAMAccountName, department, displayName, mail))
                        else:
                            self.member_table.setItem(i, 0, self.make_center_item(""))
                            self.member_table.setItem(i, 1, self.make_center_item(""))
                            self.member_table.setItem(i, 2, self.make_center_item(""))
                            self.member_table.setItem(i, 3, self.make_center_item(""))
                else:
                    self.result_text.setPlainText("그룹을 찾을 수 없습니다.")
        except ldap3.core.exceptions.LDAPBindError as e:
            QMessageBox.critical(self, "AD 인증 실패", "AD 서버 인증에 실패했습니다.\nID 또는 비밀번호를 확인해 주세요.")
            return False
        except ldap3.LDAPException as e:
            QMessageBox.critical(self, "AD 연결 오류", f"AD 서버 연결 중 오류 발생:\n{str(e)}")
            return False
        except Exception as e:
            QMessageBox.critical(self, "오류", f"오류 발생: {str(e)}")
            return False
        finally:
            self.member_table.setSortingEnabled(True)

        self.member_table.sortItems(0, Qt.AscendingOrder)
        return True

    def copy_all_members(self):
        try:
            text = ""
            for row in range(self.member_table.rowCount()):
                sAMAccountName_item = self.member_table.item(row, 0)
                department_item = self.member_table.item(row, 1)
                displayName_item = self.member_table.item(row, 2)
                mail_item = self.member_table.item(row, 3)
                sAMAccountName = sAMAccountName_item.text() if sAMAccountName_item else ""
                department = department_item.text() if department_item else ""
                displayName = displayName_item.text() if displayName_item else ""
                mail = mail_item.text() if mail_item else ""
                text += f"{sAMAccountName}\t{department}\t{displayName}\t{mail}\n"
            clipboard = QApplication.clipboard()
            clipboard.setText(text, mode=clipboard.Clipboard)
            self.result_text.setPlainText("모든 멤버 정보가 클립보드에 복사되었습니다.")
        except Exception as e:
            self.result_text.setPlainText(f"복사 중 오류가 발생했습니다: {str(e)}")

class AccountManagementDialog(QDialog):
    def __init__(self, current_server_ip, current_user, current_password):
        super().__init__()
        self.setWindowTitle("AD 서버 정보 관리")
        self.setGeometry(100, 100, 300, 100)
        
        self.server_ip_input = QLineEdit(current_server_ip)
        self.user_input = QLineEdit(current_user)
        self.password_input = QLineEdit(current_password)
        self.password_input.setEchoMode(QLineEdit.Password)
        
        self.layout = QFormLayout()
        self.layout.addRow("AD서버 IP :", self.server_ip_input)
        self.layout.addRow("관리자 ID  :", self.user_input)
        self.layout.addRow("비밀 번호  :", self.password_input)
        
        self.save_button = QPushButton("저장")
        self.save_button.clicked.connect(self.accept)
        self.layout.addWidget(self.save_button)
        
        self.setLayout(self.layout)
        self.apply_custom_styles()

        self.center()

    def center(self):
        screen = QApplication.primaryScreen()
        if screen:
            screen_geometry = screen.availableGeometry()
            self.move(screen_geometry.center() - self.rect().center())

    def get_account_info(self):
        return self.server_ip_input.text(), self.user_input.text(), self.password_input.text()
    
    def apply_custom_styles(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: #e0e0e0;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
            }

            QLineEdit {
                background-color: #3b3b3b;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 5px;
                color: #f0f0f0;
            }

            QPushButton {
                background-color: #0078d7;
                color: white;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #005a9e;
            }

            QLabel {
                font-size: 13px;
                font-weight: bold;
                color: #f0f0f0;
            }
        """)

class AccountSelectionDialog(QDialog):
    def __init__(self, entries, parent=None):
        super().__init__(parent)
        self.setWindowTitle("계정 선택")
        self.entries = entries
        self.layout = QVBoxLayout()
        self.label = QLabel("여러 계정이 검색되었습니다. 선택해 주세요:")
        self.combo = QComboBox()

        for entry in entries:
            sAMAccountName = entry["sAMAccountName"].value if "sAMAccountName" in entry else ""
            displayName = entry["displayName"].value if "displayName" in entry else ""
            mail = entry["mail"].value if "mail" in entry else ""
            self.combo.addItem(f"{displayName} ({sAMAccountName}, {mail})", entry.entry_dn)

        self.ok_button = QPushButton("확인")
        self.cancel_button = QPushButton("취소")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.ok_button)
        btn_layout.addWidget(self.cancel_button)

        self.layout.addWidget(self.label)
        self.layout.addWidget(self.combo)
        self.layout.addLayout(btn_layout)
        self.setLayout(self.layout)
        self.apply_custom_styles()

    def apply_custom_styles(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: #e0e0e0;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
            }

            QComboBox {
                background-color: #3b3b3b;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 5px;
                color: #f0f0f0;
            }

            QPushButton {
                background-color: #0078d7;
                color: white;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #005a9e;
            }

            QLabel {
                font-size: 13px;
                font-weight: bold;
                color: #f0f0f0;
            }
        """)

    def get_selected_account(self):
        index = self.combo.currentIndex()
        return self.combo.itemData(index)

class MemberManagementDialog(QDialog):
    def __init__(self, group_name, members, account_info):
        super().__init__()
        self.group_name = group_name
        self.members = members
        self.account_info = account_info
        self.setWindowTitle(f"멤버 관리 - {group_name}")
        self.resize(400, 150)
        self.layout = QVBoxLayout()
        self.apply_custom_styles()

        self.is_manual_input = False

        self.info_label = QLabel("1. 사용자 이름, 사원번호, 이메일 중 입력 \n2. 여러 명인 경우 콤마(,)로 구분")
        self.info_label.setStyleSheet("font-size: 12px; color: gray;")

        self.member_combo_label = QLabel("사용자:")
        self.member_combo = QComboBox()
        self.member_combo.setEditable(True)
        self.member_combo.setInsertPolicy(QComboBox.NoInsert)
        self.member_combo.addItem("")
        self.member_combo.addItems([f"{member.displayName} ({member.sAMAccountName}, {member.mail})" for member in members])

        completer = QCompleter()
        completer.setCompletionMode(QCompleter.CompletionMode(0))
        self.member_combo.setCompleter(completer)

        self.member_combo.lineEdit().textEdited.connect(self.handle_manual_input)

        self.member_combo.activated.connect(self.populate_display_name)

        self.browse_button = QPushButton("찾아보기")
        self.add_member_button = QPushButton("멤버 추가")
        self.remove_member_button = QPushButton("멤버 제거")
        self.add_member_button.setMinimumHeight(36)
        self.remove_member_button.setMinimumHeight(36)
        self.browse_button.setMinimumHeight(36)
        self.browse_button.setMinimumWidth(110)
        self.add_member_button.setMinimumWidth(110)
        self.remove_member_button.setMinimumWidth(110)

        self.layout.addWidget(self.info_label)
        self.layout.addWidget(self.member_combo_label)
        self.layout.addWidget(self.member_combo)
        button_row = QHBoxLayout()
        button_row.addWidget(self.browse_button, 1)
        button_row.addWidget(self.add_member_button, 1)
        button_row.addWidget(self.remove_member_button, 1)
        self.layout.addLayout(button_row)

        self.browse_button.clicked.connect(self.open_member_browse)
        self.add_member_button.clicked.connect(lambda: self.add_member())
        self.remove_member_button.clicked.connect(lambda: self.remove_member())
        self.setLayout(self.layout)
        self.progress_dialog = None

    def apply_custom_styles(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: #e0e0e0;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
            }

            QComboBox {
                background-color: #3b3b3b;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 5px;
                color: #f0f0f0;
            }

            QPushButton {
                background-color: #0078d7;
                color: white;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #005a9e;
            }

            QLabel {
                font-size: 13px;
                font-weight: bold;
                color: #f0f0f0;
            }
        """)

    def handle_manual_input(self, text):
        """사용자가 직접 입력 시 자동 채우기 비활성화"""
        self.is_manual_input = True

    def populate_display_name(self):
        """콤보박스에서 항목을 선택했을 때만 자동으로 입력 필드 채우기"""
        if not self.is_manual_input:
            selected_text = self.member_combo.currentText()
            if selected_text:
                match = re.match(r'^(.*?) \(', selected_text)
                if match:
                    display_name = match.group(1).strip()
                    self.member_combo.setEditText(display_name)
        else:
            self.is_manual_input = False

    def open_member_browse(self):
        dialog = MemberBrowseDialog(self.account_info, self)
        if dialog.exec_() == QDialog.Accepted:
            selected_ids = dialog.get_selected_ids()
            if selected_ids:
                existing_text = self.member_combo.currentText().strip()
                if existing_text:
                    merged = [existing_text] + selected_ids
                else:
                    merged = selected_ids
                self.member_combo.setEditText(", ".join(merged))

    def run_powershell_command(self, command):
        try:
            subprocess.run(
                ["powershell", "-Command", command],
                check=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            QMessageBox.information(self, "성공", "성공적으로 실행되었습니다.")
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "오류", f"실행에 실패했습니다.\n{str(e)}")

    def get_group_dn(self, conn):
        escaped_group_name = escape_filter_chars(self.group_name)
        search_filter = f"(|(cn={escaped_group_name})(displayName={escaped_group_name}))"
        conn.search(
            search_base="DC=lskglobal,DC=com",
            search_filter=search_filter,
            attributes=["distinguishedName"]
        )
        if not conn.entries:
            return None
        return conn.entries[0].entry_dn

    def resolve_identifier(self, identifier, conn):
        escaped_identifier = escape_filter_chars(identifier)
        search_filter = (
            f"(|"
            f"(displayName=*{escaped_identifier}*)"
            f"(sAMAccountName=*{escaped_identifier}*)"
            f"(mail=*{escaped_identifier}*)"
            f")"
        )

        try:
            conn.search(
                search_base="DC=lskglobal,DC=com",
                search_filter=search_filter,
                attributes=["sAMAccountName", "displayName", "mail", "distinguishedName"]
            )
        except LDAPInvalidFilterError as e:
            QMessageBox.critical(
                self,
                "LDAP 오류",
                f"잘못된 필터: {search_filter}\n{str(e)}"
            )
            return None

        if conn.entries:
            if len(conn.entries) == 1:
                return conn.entries[0].entry_dn
            dialog = AccountSelectionDialog(conn.entries, self)
            if dialog.exec_() == QDialog.Accepted:
                selected_dn = dialog.get_selected_account()
                return selected_dn
            else:
                QMessageBox.information(
                    self,
                    "취소",
                    "계정 선택이 취소되었습니다."
                )
            return None

        QMessageBox.critical(
            self,
            "오류",
            "사용자를 찾을 수 없습니다."
        )
        return None

    def add_member(self):
        identifiers = self.member_combo.currentText().strip().split(',')
        identifiers = [identifier.strip() for identifier in identifiers if identifier.strip()]

        if not identifiers:
            QMessageBox.warning(self, "경고", "추가할 사용자를 입력하세요.")
            return

        server_uri = f"ldap://{self.account_info['server_ip']}"
        try:
            with ldap3.Connection(
                server_uri,
                user=self.account_info['user'],
                password=self.account_info['password'],
                auto_bind=True
            ) as conn:
                group_dn = self.get_group_dn(conn)
                if not group_dn:
                    QMessageBox.critical(self, "오류", f"그룹 '{self.group_name}' DN을 찾을 수 없습니다.")
                    return

                resolved_list = []
                for identifier in identifiers:
                    user_dn = self.resolve_identifier(identifier, conn)
                    if user_dn is not None:
                        resolved_list.append((identifier, user_dn))

                if not resolved_list:
                    QMessageBox.warning(self, "경고", "입력한 정보로 찾은 사용자가 없습니다.")
                    return

                self.progress_dialog = QProgressDialog("사용자 추가 중...", "취소", 0, len(resolved_list), self)
                self.progress_dialog.setWindowTitle("진행 상황")
                self.progress_dialog.setWindowModality(Qt.WindowModal)
                self.progress_dialog.setMinimumDuration(0)

                success_count = 0
                failed_identifiers = []
                for index, (orig_identifier, user_dn) in enumerate(resolved_list, start=1):
                    if self.progress_dialog.wasCanceled():
                        QMessageBox.information(self, "알림", "작업이 취소되었습니다.")
                        return

                    modify_result = conn.modify(
                        group_dn,
                        {"member": [(ldap3.MODIFY_ADD, [user_dn])]}
                    )
                    if modify_result:
                        success_count += 1
                    else:
                        failed_identifiers.append(f"{orig_identifier}: {conn.result}")

                    self.progress_dialog.setValue(index)
                    self.progress_dialog.setLabelText(f"사용자 {index}/{len(resolved_list)} 추가 중...")
                    QApplication.processEvents()

                self.progress_dialog.close()
                if failed_identifiers:
                    QMessageBox.warning(self, "일부 실패", "다음 사용자 추가에 실패했습니다:\n" + "\n".join(failed_identifiers))
                QMessageBox.information(self, "완료", f"{success_count}/{len(resolved_list)}명의 사용자 추가 작업이 완료되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"LDAP 서버에 연결할 수 없습니다.\n오류: {str(e)}")

    def remove_member(self):
        identifiers = self.member_combo.currentText().strip().split(',')
        identifiers = [identifier.strip() for identifier in identifiers if identifier.strip()]

        if not identifiers:
            QMessageBox.warning(self, "경고", "제거할 사용자를 입력하세요.")
            return

        server_uri = f"ldap://{self.account_info['server_ip']}"
        try:
            with ldap3.Connection(
                server_uri,
                user=self.account_info['user'],
                password=self.account_info['password'],
                auto_bind=True
            ) as conn:
                group_dn = self.get_group_dn(conn)
                if not group_dn:
                    QMessageBox.critical(self, "오류", f"그룹 '{self.group_name}' DN을 찾을 수 없습니다.")
                    return

                resolved_list = []
                for identifier in identifiers:
                    user_dn = self.resolve_identifier(identifier, conn)
                    if user_dn is not None:
                        resolved_list.append((identifier, user_dn))

                if not resolved_list:
                    QMessageBox.warning(self, "경고", "입력한 정보로 찾은 사용자가 없습니다.")
                    return

                self.progress_dialog = QProgressDialog("사용자 제거 중...", "취소", 0, len(resolved_list), self)
                self.progress_dialog.setWindowTitle("진행 상황")
                self.progress_dialog.setWindowModality(Qt.WindowModal)
                self.progress_dialog.setMinimumDuration(0)

                success_count = 0
                failed_identifiers = []
                for index, (orig_identifier, user_dn) in enumerate(resolved_list, start=1):
                    if self.progress_dialog.wasCanceled():
                        QMessageBox.information(self, "알림", "작업이 취소되었습니다.")
                        return

                    modify_result = conn.modify(
                        group_dn,
                        {"member": [(ldap3.MODIFY_DELETE, [user_dn])]}
                    )
                    if modify_result:
                        success_count += 1
                    else:
                        failed_identifiers.append(f"{orig_identifier}: {conn.result}")

                    self.progress_dialog.setValue(index)
                    self.progress_dialog.setLabelText(f"사용자 {index}/{len(resolved_list)} 제거 중...")
                    QApplication.processEvents()

                self.progress_dialog.close()
                if failed_identifiers:
                    QMessageBox.warning(self, "일부 실패", "다음 사용자 제거에 실패했습니다:\n" + "\n".join(failed_identifiers))
                QMessageBox.information(self, "완료", f"{success_count}/{len(resolved_list)}명의 사용자 제거 작업이 완료되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"LDAP 서버에 연결할 수 없습니다.\n오류: {str(e)}")


class MemberBrowseDialog(QDialog):
    def __init__(self, account_info, parent=None):
        super().__init__(parent)
        self.account_info = account_info
        self.selected_ids = []
        self.setWindowTitle("사용자 찾아보기")
        self.resize(900, 600)

        self.layout = QVBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("사원번호 / 이름 / 메일 / 부서 검색")
        self.search_input.setClearButtonEnabled(True)
        self.tag_scroll = QScrollArea()
        self.tag_scroll.setWidgetResizable(True)
        self.tag_scroll.setFixedHeight(56)
        self.tag_widget = QWidget()
        self.tag_layout = QHBoxLayout()
        self.tag_layout.setContentsMargins(6, 6, 6, 6)
        self.tag_layout.setSpacing(6)
        self.tag_widget.setLayout(self.tag_layout)
        self.tag_scroll.setWidget(self.tag_widget)
        self.user_table = QTableWidget()
        self.user_table.setColumnCount(4)
        self.user_table.setHorizontalHeaderLabels(["사원 번호", "부서", "표시 이름", "메일 주소"])
        self.user_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.user_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.user_table.setSelectionMode(QAbstractItemView.MultiSelection)
        self.user_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.user_table.verticalHeader().setVisible(False)
        self.user_table.setSortingEnabled(True)

        button_layout = QHBoxLayout()
        self.select_button = QPushButton("선택")
        self.cancel_button = QPushButton("취소")
        button_layout.addStretch()
        button_layout.addWidget(self.select_button)
        button_layout.addWidget(self.cancel_button)

        self.layout.addWidget(QLabel("사용자 검색"))
        self.layout.addWidget(self.tag_scroll)
        self.layout.addWidget(self.search_input)
        self.layout.addWidget(self.user_table)
        self.layout.addLayout(button_layout)
        self.setLayout(self.layout)

        self.search_input.textChanged.connect(self.filter_table)
        self.search_input.textEdited.connect(self.filter_table)
        self.search_input.returnPressed.connect(lambda: self.filter_table(self.search_input.text()))
        self.search_input.installEventFilter(self)
        self.user_table.itemSelectionChanged.connect(self.update_selected_tags)
        self.select_button.clicked.connect(self.accept_selection)
        self.cancel_button.clicked.connect(self.reject)
        self.select_button.setAutoDefault(False)
        self.select_button.setDefault(False)
        self.cancel_button.setAutoDefault(False)
        self.cancel_button.setDefault(False)
        self.last_search_text = ""
        self.search_poll_timer = QtCore.QTimer(self)
        self.search_poll_timer.setInterval(120)
        self.search_poll_timer.timeout.connect(self.poll_search_text)
        self.search_poll_timer.start()
        self.load_users()
        self.update_selected_tags()

    def make_center_item(self, value):
        item = QTableWidgetItem(str(value) if value is not None else "")
        item.setTextAlignment(Qt.AlignCenter)
        return item

    def load_users(self):
        server_uri = f"ldap://{self.account_info['server_ip']}"
        try:
            with ldap3.Connection(
                server_uri,
                user=self.account_info['user'],
                password=self.account_info['password'],
                auto_bind=True
            ) as conn:
                conn.search(
                    search_base="OU=lskglobal,DC=lskglobal,DC=com",
                    search_filter="(&(objectClass=user)(objectCategory=person))",
                    attributes=["sAMAccountName", "department", "displayName", "mail"],
                    paged_size=500
                )
                entries = conn.entries

            self.user_table.setSortingEnabled(False)
            self.user_table.setRowCount(len(entries))
            for i, entry in enumerate(entries):
                sAMAccountName = entry["sAMAccountName"].value if "sAMAccountName" in entry else ""
                department = entry["department"].value if "department" in entry else ""
                displayName = entry["displayName"].value if "displayName" in entry else ""
                mail = entry["mail"].value if "mail" in entry else ""

                self.user_table.setItem(i, 0, self.make_center_item(sAMAccountName))
                self.user_table.setItem(i, 1, self.make_center_item(department))
                self.user_table.setItem(i, 2, self.make_center_item(displayName))
                self.user_table.setItem(i, 3, self.make_center_item(mail))
            self.user_table.setSortingEnabled(True)
            self.user_table.sortItems(0, Qt.AscendingOrder)
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"사용자 목록 조회 실패:\n{str(e)}")

    def filter_table(self, text):
        keyword = text.strip().lower()
        for row in range(self.user_table.rowCount()):
            show = False
            for col in range(self.user_table.columnCount()):
                item = self.user_table.item(row, col)
                if item:
                    item_text = item.text().lower()
                    initials = self.extract_korean_initials(item_text)
                    if keyword in item_text or (keyword and keyword in initials):
                        show = True
                        break
            self.user_table.setRowHidden(row, not show)

    def eventFilter(self, obj, event):
        if obj is self.search_input and event.type() == QtCore.QEvent.KeyPress:
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self.filter_table(self.search_input.text())
                return True
        if obj is self.search_input and event.type() in (QtCore.QEvent.KeyRelease, QtCore.QEvent.InputMethod):
            QtCore.QTimer.singleShot(0, lambda: self.filter_table(self.search_input.text()))
        return super().eventFilter(obj, event)

    def extract_korean_initials(self, text):
        choseong = [
            "ㄱ", "ㄲ", "ㄴ", "ㄷ", "ㄸ", "ㄹ", "ㅁ", "ㅂ", "ㅃ", "ㅅ",
            "ㅆ", "ㅇ", "ㅈ", "ㅉ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"
        ]
        result = []
        for ch in text:
            code = ord(ch)
            if 0xAC00 <= code <= 0xD7A3:
                index = (code - 0xAC00) // 588
                result.append(choseong[index])
            else:
                result.append(ch)
        return "".join(result)

    def accept_selection(self):
        rows = self.user_table.selectionModel().selectedRows()
        self.selected_ids = []
        for row_idx in rows:
            row = row_idx.row()
            id_item = self.user_table.item(row, 0)
            if id_item and id_item.text().strip():
                self.selected_ids.append(id_item.text().strip())
        self.accept()

    def get_selected_ids(self):
        return self.selected_ids

    def poll_search_text(self):
        current_text = self.search_input.text()
        if current_text != self.last_search_text:
            self.last_search_text = current_text
            self.filter_table(current_text)

    def update_selected_tags(self):
        while self.tag_layout.count():
            item = self.tag_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        selected_rows = self.user_table.selectionModel().selectedRows()
        if not selected_rows:
            placeholder = QLabel("선택된 사용자가 여기에 표시됩니다.")
            placeholder.setStyleSheet("color: #aaaaaa; padding: 4px;")
            self.tag_layout.addWidget(placeholder)
            self.tag_layout.addStretch()
            return

        for row_idx in selected_rows:
            row = row_idx.row()
            emp = self.user_table.item(row, 0).text() if self.user_table.item(row, 0) else ""
            name = self.user_table.item(row, 2).text() if self.user_table.item(row, 2) else ""
            tag = QLabel(f"{name} ({emp})")
            tag.setStyleSheet(
                "background-color:#3d7eff; color:white; border-radius:10px; "
                "padding:4px 10px; font-size:11px;"
            )
            self.tag_layout.addWidget(tag)
        self.tag_layout.addStretch()


class AddParentGroupsDialog(QDialog):
    def __init__(self, account_info, parent=None):
        super().__init__(parent)
        self.account_info = account_info
        self.setWindowTitle("소속 그룹 추가")
        self.resize(480, 180)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("추가할 그룹 메일 주소(여러 개는 콤마로 구분):"))
        self.mail_input = QLineEdit()
        self.mail_input.setPlaceholderText("group1@..., group2@...")
        layout.addWidget(self.mail_input)

        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("추가")
        self.cancel_button = QPushButton("취소")
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)
        self.setLayout(layout)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

    def get_mails(self):
        raw = self.mail_input.text().strip()
        return [v.strip() for v in raw.split(",") if v.strip()]


class RemoveParentGroupsDialog(QDialog):
    def __init__(self, groups, parent=None):
        super().__init__(parent)
        self.setWindowTitle("소속 그룹 삭제")
        self.resize(520, 420)
        self.groups = groups

        layout = QVBoxLayout()
        layout.addWidget(QLabel("삭제할 소속 그룹을 선택하세요(다중 선택 가능):"))
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        for group in groups:
            item = QListWidgetItem(group.get("display", group.get("dn", "")))
            item.setData(Qt.UserRole, group.get("dn", ""))
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("삭제")
        self.cancel_button = QPushButton("취소")
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)
        self.setLayout(layout)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

    def get_selected_dns(self):
        items = self.list_widget.selectedItems()
        return [item.data(Qt.UserRole) for item in items]


class GroupManagementDialog(QDialog):
    def __init__(self, group_name, account_info, parent=None):
        super().__init__(parent)
        self.group_name = group_name
        self.account_info = account_info
        self.group_dn = ""
        self.parent_groups = []

        self.setWindowTitle(f"그룹 관리 - {group_name}")
        self.resize(700, 520)

        self.layout = QVBoxLayout()
        self.group_info_label = QLabel("그룹 설명")
        self.desc_input = QLineEdit()
        self.save_desc_button = QPushButton("설명 저장")

        self.parent_group_label = QLabel("소속 그룹")
        self.parent_group_list = QListWidget()
        self.parent_group_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.parent_group_list.setMinimumHeight(260)
        self.add_parent_button = QPushButton("소속 그룹 추가")
        self.remove_parent_button = QPushButton("소속 그룹 삭제")

        desc_layout = QHBoxLayout()
        desc_layout.addWidget(self.desc_input)
        desc_layout.addWidget(self.save_desc_button)

        parent_btn_layout = QHBoxLayout()
        parent_btn_layout.addWidget(self.add_parent_button)
        parent_btn_layout.addWidget(self.remove_parent_button)
        parent_btn_layout.addStretch()

        self.layout.addWidget(self.group_info_label)
        self.layout.addLayout(desc_layout)
        self.layout.addSpacing(10)
        self.layout.addWidget(self.parent_group_label)
        self.layout.addWidget(self.parent_group_list)
        self.layout.addLayout(parent_btn_layout)
        self.setLayout(self.layout)

        self.save_desc_button.clicked.connect(self.save_description)
        self.add_parent_button.clicked.connect(self.add_parent_groups)
        self.remove_parent_button.clicked.connect(self.remove_parent_groups)

        self.load_group_info()

    def get_connection(self):
        server_uri = f"ldap://{self.account_info['server_ip']}"
        return ldap3.Connection(
            server_uri,
            user=self.account_info['user'],
            password=self.account_info['password'],
            auto_bind=True
        )

    def load_group_info(self):
        try:
            with self.get_connection() as conn:
                filter_group = escape_filter_chars(self.group_name)
                conn.search(
                    search_base="DC=lskglobal,DC=com",
                    search_filter=f"(|(cn={filter_group})(displayName={filter_group}))",
                    attributes=["distinguishedName", "description", "memberOf", "cn", "mail"]
                )
                if not conn.entries:
                    QMessageBox.warning(self, "경고", "그룹을 찾을 수 없습니다.")
                    self.reject()
                    return

                group_entry = conn.entries[0]
                self.group_dn = group_entry.entry_dn
                description = group_entry["description"].value if "description" in group_entry else ""
                self.desc_input.setText(description if description else "")

                member_of_dns = [str(v) for v in group_entry["memberOf"]] if "memberOf" in group_entry else []
                self.parent_groups = []
                self.parent_group_list.clear()
                if member_of_dns:
                    dn_filter = "".join(f"(distinguishedName={escape_filter_chars(dn)})" for dn in member_of_dns)
                    conn.search(
                        search_base="DC=lskglobal,DC=com",
                        search_filter=f"(|{dn_filter})",
                        attributes=["cn", "mail", "distinguishedName"]
                    )
                    detail_map = {entry.entry_dn: entry for entry in conn.entries}
                    for parent_dn in member_of_dns:
                        info = detail_map.get(parent_dn)
                        cn = info["cn"].value if info and "cn" in info else parent_dn
                        mail = info["mail"].value if info and "mail" in info else ""
                        display = f"{cn} ({mail})" if mail else cn
                        self.parent_groups.append({"dn": parent_dn, "display": display})
                        self.parent_group_list.addItem(display)
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"그룹 정보 조회 실패:\n{str(e)}")
            self.reject()

    def save_description(self):
        if not self.group_dn:
            return
        new_desc = self.desc_input.text().strip()
        try:
            with self.get_connection() as conn:
                conn.modify(self.group_dn, {"description": [(ldap3.MODIFY_REPLACE, [new_desc])]})
                if conn.result.get("result") == 0:
                    QMessageBox.information(self, "성공", "그룹 설명이 저장되었습니다.")
                else:
                    QMessageBox.critical(self, "실패", f"설명 저장 실패:\n{conn.result}")
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"설명 저장 실패:\n{str(e)}")

    def add_parent_groups(self):
        dialog = AddParentGroupsDialog(self.account_info, self)
        if dialog.exec_() != QDialog.Accepted:
            return

        mails = dialog.get_mails()
        if not mails:
            QMessageBox.warning(self, "경고", "추가할 메일 주소를 입력하세요.")
            return

        failed = []
        success = 0
        try:
            with self.get_connection() as conn:
                for mail in mails:
                    escaped_mail = escape_filter_chars(mail)
                    conn.search(
                        search_base="DC=lskglobal,DC=com",
                        search_filter=f"(mail={escaped_mail})",
                        attributes=["distinguishedName", "cn"]
                    )
                    if not conn.entries:
                        failed.append(f"{mail}: 메일 주소와 일치하는 그룹 없음")
                        continue

                    parent_dn = conn.entries[0].entry_dn
                    result = conn.modify(parent_dn, {"member": [(ldap3.MODIFY_ADD, [self.group_dn])]})
                    if result:
                        success += 1
                    else:
                        failed.append(f"{mail}: {conn.result}")

            self.load_group_info()
            if failed:
                QMessageBox.warning(self, "일부 실패", "\n".join(failed))
            QMessageBox.information(self, "완료", f"{success}/{len(mails)}개 소속 그룹 추가 완료")
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"소속 그룹 추가 실패:\n{str(e)}")

    def remove_parent_groups(self):
        if not self.parent_groups:
            QMessageBox.information(self, "알림", "삭제할 소속 그룹이 없습니다.")
            return

        dialog = RemoveParentGroupsDialog(self.parent_groups, self)
        if dialog.exec_() != QDialog.Accepted:
            return

        selected_dns = dialog.get_selected_dns()
        if not selected_dns:
            QMessageBox.warning(self, "경고", "삭제할 항목을 선택하세요.")
            return

        failed = []
        success = 0
        try:
            with self.get_connection() as conn:
                for parent_dn in selected_dns:
                    result = conn.modify(parent_dn, {"member": [(ldap3.MODIFY_DELETE, [self.group_dn])]})
                    if result:
                        success += 1
                    else:
                        failed.append(f"{parent_dn}: {conn.result}")

            self.load_group_info()
            if failed:
                QMessageBox.warning(self, "일부 실패", "\n".join(failed))
            QMessageBox.information(self, "완료", f"{success}/{len(selected_dns)}개 소속 그룹 삭제 완료")
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"소속 그룹 삭제 실패:\n{str(e)}")


class CreateGroupDialog(QDialog):
    def __init__(self, group_name="", account_info=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("새 그룹 생성")
        self.setGeometry(100, 100, 350, 200)

        self.account_info = account_info
        self.layout = QVBoxLayout()

        self.name_label = QLabel("그룹 이름:")
        self.name_input = QLineEdit()
        self.name_input.setText(group_name)

        self.desc_label = QLabel("그룹 설명:")
        self.desc_input = QLineEdit()

        self.security_group_button = QPushButton("보안 그룹")
        self.mail_group_button = QPushButton("그룹 메일")

        self.scope_label = QLabel("그룹 범위:")
        self.scope_combo = QComboBox()
        self.scope_combo.addItems(["도메인 로컬", "글로벌", "유니버설"])

        self.type_label = QLabel("그룹 종류:")
        self.type_combo = QComboBox()
        self.type_combo.addItems(["보안", "배포"])

        self.create_button = QPushButton("생성")
        self.cancel_button = QPushButton("취소")

        self.create_button.clicked.connect(self.create_group)
        self.cancel_button.clicked.connect(self.reject)
        
        self.security_group_button.clicked.connect(lambda: self.set_group_type("도메인 로컬", "보안"))
        self.mail_group_button.clicked.connect(lambda: self.set_group_type("유니버설", "배포"))

        form_layout = QFormLayout()
        form_layout.addRow(self.name_label, self.name_input)
        form_layout.addRow(self.desc_label, self.desc_input)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.security_group_button)
        button_layout.addWidget(self.mail_group_button)

        detail_layout = QFormLayout()
        detail_layout.addRow(self.scope_label, self.scope_combo)
        detail_layout.addRow(self.type_label, self.type_combo)

        bottom_button_layout = QHBoxLayout()
        bottom_button_layout.addWidget(self.create_button)
        bottom_button_layout.addWidget(self.cancel_button)

        self.layout.addLayout(form_layout)
        self.layout.addLayout(button_layout)
        self.layout.addLayout(detail_layout)
        self.layout.addLayout(bottom_button_layout)

        self.setLayout(self.layout)
        self.center()

    def center(self):
        screen = QApplication.primaryScreen()
        if screen:
            screen_geometry = screen.availableGeometry()
            self.move(screen_geometry.center() - self.rect().center())

    def set_group_type(self, scope, group_type):
        self.scope_combo.setCurrentText(scope)
        self.type_combo.setCurrentText(group_type)

    def create_group(self):
        group_name = self.name_input.text().strip()
        group_desc = self.desc_input.text().strip()
        group_scope = self.scope_combo.currentText()
        group_type = self.type_combo.currentText()

        server_ip = self.account_info.get('server_ip', '').strip()
        if not server_ip:
            QMessageBox.critical(self, "오류", "AD 서버 IP가 입력되지 않았습니다.\n관리자 계정 관리에서 설정해 주세요.")
            return

        if not group_name:
            QMessageBox.warning(self, "경고", "그룹 이름을 입력하세요.")
            return

        if self.check_group_exists(group_name):
            QMessageBox.warning(self, "경고", f"'{group_name}' 그룹은 이미 존재합니다.")
            return

        scope_mapping = {"도메인 로컬": "DomainLocal", "글로벌": "Global", "유니버설": "Universal"}
        type_mapping = {"보안": "Security", "배포": "Distribution"}

        scope_english = scope_mapping.get(group_scope, "Global")
        type_english = type_mapping.get(group_type, "Security")

        if scope_english == "Universal" and type_english == "Distribution":
            group_name = group_name.lower()

        command = (
            f"New-ADGroup -Name '{group_name}' "
            f"-SamAccountName '{group_name}' "
            f"-GroupScope '{scope_english}' -GroupCategory '{type_english}' "
            f"-Description '{group_desc}' -Path 'OU=Users,OU=lskglobal,DC=lskglobal,DC=com'"
        )

        if type_english == "Distribution" and scope_english == "Universal":
            command += f" -DisplayName '{group_name}' -OtherAttributes @{{'mail'='{group_name}'}}"
            command += f"; Set-ADGroup -Identity '{group_name}' -add @{{'msExchHideFromAddressLists'=$true; 'msExchRequireAuthToSendTo'=$False}}"

        try:
            subprocess.run(
                ["powershell", "-Command", command],
                check=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            self.accept()
            
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "오류", f"그룹 생성 중 오류 발생:\n{str(e)}")

    def check_group_exists(self, group_name):
        server_ip = self.account_info.get('server_ip', '').strip()
        if not server_ip:
            QMessageBox.critical(self, "오류", "AD 서버 정보가 설정되지 않았습니다.")
            return False

        server_uri = f"ldap://{server_ip}"
        escaped_group_name = escape_filter_chars(group_name)
        search_filter = f"(cn={escaped_group_name})"

        try:
            with ldap3.Connection(server_uri, user=self.account_info['user'], password=self.account_info['password'], auto_bind=True) as conn:
                conn.search(search_base="DC=lskglobal,DC=com", search_filter=search_filter, attributes=["cn"])
                return len(conn.entries) > 0
        except Exception as e:
            QMessageBox.critical(self, "LDAP 오류", f"AD 조회 중 오류 발생:\n{str(e)}")
            return False

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ADGroupViewer()
    ex.show()
    sys.exit(app.exec_())
