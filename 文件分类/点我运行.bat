@echo off
setlocal enabledelayedexpansion
chcp 65001 > nul

rem 获取脚本所在的文件夹路径
set "source_folder=%~dp0"

rem 设置目标文件夹路径
set "ppt_folder=%source_folder%PPT_Files"
set "doc_folder=%source_folder%DOC_Files"
set "media_folder=%source_folder%Media_Files"

rem 检查桌面是否有符合要求的文件
set "ppt_exist=false"
set "doc_exist=false"
set "media_exist=false"

for %%i in ("%source_folder%\*") do (
    set "ext=%%~xi"
    if /i "!ext!"==".ppt" set "ppt_exist=true"
    if /i "!ext!"==".pptx" set "ppt_exist=true"
    if /i "!ext!"==".doc" set "doc_exist=true"
    if /i "!ext!"==".docx" set "doc_exist=true"
    if /i "!ext!"==".jpg" set "media_exist=true"
    if /i "!ext!"==".png" set "media_exist=true"
    if /i "!ext!"==".mp4" set "media_exist=true"
    if /i "!ext!"==".avi" set "media_exist=true"
)

rem 创建目标文件夹（如果存在对应类型的文件）
if "%ppt_exist%"=="true" mkdir "%ppt_folder%"
if "%doc_exist%"=="true" mkdir "%doc_folder%"
if "%media_exist%"=="true" mkdir "%media_folder%"

rem 遍历源文件夹中的文件
for %%i in ("%source_folder%\*") do (
    set "ext=%%~xi"
    rem 移动文件到相应的文件夹
    if /i "!ext!"==".ppt" (
        set "target_path=%ppt_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    if /i "!ext!"==".pptx" (
        set "target_path=%ppt_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    if /i "!ext!"==".doc" (
        set "target_path=%doc_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    if /i "!ext!"==".docx" (
        set "target_path=%doc_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    rem 将视频和图片归类到一个文件夹
    if /i "!ext!"==".jpg" (
        set "target_path=%media_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    if /i "!ext!"==".png" (
        set "target_path=%media_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    if /i "!ext!"==".mp4" (
        set "target_path=%media_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    if /i "!ext!"==".avi" (
        set "target_path=%media_folder%\%%~nxi"
        call :move_with_rename "%%i" "!target_path!"
    )
    rem 添加其他文件类型的归类规则
)

rem 输出任务完成信息
echo 任务完成！
pause
exit /b

:move_with_rename
rem 移动文件到目标文件夹并重命名以避免覆盖
set "source_file=%~1"
set "target_file=%~2"
if exist "%target_file%" (
    set "index=1"
    :rename_loop
    if exist "%target_file%_!index!" (
        set /a "index+=1"
        goto :rename_loop
    )
    set "target_file=%target_file%_!index!"
)
move "%source_file%" "%target_file%"
echo 文件 "%source_file%" 已移动到 "%target_file%"
exit /b
