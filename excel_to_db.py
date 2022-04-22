from .base import DjangoExcelToDB
from materials.models import Material


class StockWorkBook(DjangoExcelToDB):
    model = Material
    

if __name__ == '__main__':
    stock_workbook = StockWorkBook()
    stock_workbook.migrate_to_db()

    # fields_to_extract_from_db = [
    #     'material_umesc', 'material_grp_description', 'material_class', 'material_description',
    #     'reorder_point', 'maximm_stock_level', 'total_stock', 'demand_Q'
    # ]