from ModifyTemplateExport  import GenerateExportCode
from dynaconf import settings,Validator

gcode=GenerateExportCode()

dirpath=settings.DIRPATH

gcode.batch_build(dirpath)







