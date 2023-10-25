# 폴더 생성 및 파일 이동 함수 선언 및 정의
def create_and_move_folder(file_extension, folder_name, keyword): #file_extension : 확장자 / folder_name : 폴더명 / keyword : 파일명에 포함할 키워드

    import os
    import shutil
    current_directory = os.getcwd()

    #폴더가 없으면 폴더를 만듬
    if not os.path.exists(folder_name):
        os.mkdir(folder_name)

    for filename in os.listdir(current_directory):
        if filename.endswith(file_extension) and keyword in filename:
            source_path = os.path.join(current_directory, filename)
            destination_path = os.path.join(folder_name, filename)
            shutil.move(source_path, destination_path)