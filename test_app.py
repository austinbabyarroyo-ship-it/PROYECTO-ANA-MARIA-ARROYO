import unittest
from unittest.mock import patch
import app
from io import BytesIO
from openpyxl import load_workbook

class Pruebas(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()
        self.tasas = {'USD':1,'COP':4000,'EUR':0.8}
        self.patcher = patch('app.cargar_tasas', return_value=(self.tasas,False))
        self.patcher.start()
        self.addCleanup(self.patcher.stop)
    def test_inicio(self):
        self.assertIn(b'AMVEXA', self.client.get('/').data)
    def test_tasas(self):
        self.assertEqual(self.client.get('/api/tasas').json['tasas']['COP'],4000)
    def test_pdf(self):
        r=self.client.post('/api/exportar/pdf',json={'monto':100,'de':'USD','a':'COP','comision':2})
        self.assertEqual(r.status_code,200)
        self.assertTrue(r.data.startswith(b'%PDF'))
    def test_excel(self):
        r=self.client.post('/api/exportar/xlsx',json={'monto':100,'de':'USD','a':'COP'})
        self.assertEqual(r.status_code,200)
        self.assertEqual(load_workbook(BytesIO(r.data)).active['D2'].value,400000)
    def test_invalido(self):
        self.assertEqual(self.client.post('/api/exportar/pdf',json={'monto':-1,'de':'USD','a':'COP'}).status_code,400)
    def test_historia_misma_moneda(self):
        self.assertEqual(self.client.get('/api/historia?de=USD&a=USD&dias=30').json['valores'],[1,1])
    def test_historia_invalida(self):
        self.assertEqual(self.client.get('/api/historia?de=USD&a=EUR&dias=500').status_code,400)
if __name__=='__main__': unittest.main()
