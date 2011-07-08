@echo off
@cd D:\Projects\vbasic\Configurator
@copy /Y cop.exe C:\Tools\upx303w > pack.log
@C:\Tools\upx303w\upx.exe C:\Tools\upx303w\cop.exe >> pack.log
@copy /Y C:\Tools\upx303w\cop.exe . >> pack.log
@del C:\Tools\upx303w\cop.exe >> pack.log