import base64
import sys
import requests
import os
from os.path import basename
from zipfile import ZipFile
from openerp import models, fields,api
from lxml import etree as ET
from datetime import datetime, timedelta

class common_wizard(models.TransientModel):
    _name = 'soupese.common_wizard'

    file_name = fields.Char(string='File Name', size=64)
    file =  fields.Binary('File')

    def _addField(self, fields, attrname=None, fieldtype=None, SUBTYPE=None, WIDTH=None, required=None, DECIMALS=None):
        field = ET.SubElement(fields, "FIELD")
        if attrname is not None:
            field.set("attrname", attrname)
        if fieldtype is not None:
            field.set("fieldtype", fieldtype)
        if required is not None:
            field.set("required", required)
        if DECIMALS is not None:
            field.set("DECIMALS", DECIMALS)
        if SUBTYPE is not None:
            field.set("SUBTYPE", SUBTYPE)
        if WIDTH is not None:
            field.set("WIDTH", WIDTH)
        return field

    @api.multi
    def confirm_export(self):
        filename = "SL_DO"
        files_path = '/tmp/po'
        files_pack = []
        zipfile_date = ""

        if not os.path.exists(files_path):
            os.makedirs(files_path)

        recs = self.env['purchase.order'].search([('id', 'in', self._context.get('active_ids'))])

        isBulk = True if len(recs) > 1 else False

        for rec in recs:
            po_code = ''
            supplier_code = ''
            datapacket = ET.Element("DATAPACKET", Version="2.0")
            metadata = ET.SubElement(datapacket, "METADATA")
            fields = ET.SubElement(metadata, "FIELDS")
            self._addField(fields, attrname="DOCKEY", fieldtype="i4", required="true")
            self._addField(fields, attrname="DOCNO", fieldtype="string", required="true", WIDTH="20")
            self._addField(fields, attrname="DOCNOEX", fieldtype="string", WIDTH="20")
            self._addField(fields, attrname="DOCDATE", fieldtype="date")
            self._addField(fields, attrname="POSTDATE", fieldtype="date")
            self._addField(fields, attrname="TAXDATE", fieldtype="date")
            self._addField(fields, attrname="CODE", fieldtype="string", WIDTH="10")
            self._addField(fields, attrname="COMPANYNAME", fieldtype="string", WIDTH="100")
            self._addField(fields, attrname="ADDRESS1", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="ADDRESS2", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="ADDRESS3", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="ADDRESS4", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="PHONE1", fieldtype="string", WIDTH="30")
            self._addField(fields, attrname="FAX1", fieldtype="string", WIDTH="30")
            self._addField(fields, attrname="ATTENTION", fieldtype="string", WIDTH="70")
            self._addField(fields, attrname="AREA", fieldtype="string", WIDTH="10")
            self._addField(fields, attrname="AGENT", fieldtype="string", WIDTH="10")
            self._addField(fields, attrname="PROJECT", fieldtype="string", WIDTH="20")
            self._addField(fields, attrname="TERMS", fieldtype="string", WIDTH="10")
            self._addField(fields, attrname="CURRENCYCODE", fieldtype="string", WIDTH="6")
            self._addField(fields, attrname="CURRENCYRATE", fieldtype="fixedFMT", DECIMALS="10", WIDTH="19")
            self._addField(fields, attrname="SHIPPER", fieldtype="string", required="true", WIDTH="30")
            self._addField(fields, attrname="DESCRIPTION", fieldtype="string", WIDTH="200")
            self._addField(fields, attrname="COUNTRY", fieldtype="string", WIDTH="50")
            self._addField(fields, attrname="CANCELLED", fieldtype="string", WIDTH="1")
            self._addField(fields, attrname="DOCAMT", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(fields, attrname="LOCALDOCAMT", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            # self._addField(fields, attrname="D_DOCNO", fieldtype="string", WIDTH="20")
            # self._addField(fields, attrname="D_PAYMENTMETHOD", fieldtype="string", WIDTH="10")
            # self._addField(fields, attrname="D_CHEQUENUMBER", fieldtype="string", WIDTH="20")
            # self._addField(fields, attrname="D_PAYMENTPROJECT", fieldtype="string", WIDTH="20")
            # self._addField(fields, attrname="D_BANKCHARGE", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(fields, attrname="D_AMOUNT", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(fields, attrname="VALIDITY", fieldtype="string", WIDTH="200")
            self._addField(fields, attrname="DELIVERYTERM", fieldtype="string", WIDTH="200")
            self._addField(fields, attrname="CC", fieldtype="string", WIDTH="200")
            self._addField(fields, attrname="DOCREF1", fieldtype="string", WIDTH="25")
            self._addField(fields, attrname="DOCREF2", fieldtype="string", WIDTH="25")
            self._addField(fields, attrname="DOCREF3", fieldtype="string", WIDTH="25")
            self._addField(fields, attrname="DOCREF4", fieldtype="string", WIDTH="25")
            self._addField(fields, attrname="BRANCHNAME", fieldtype="string", WIDTH="100")
            self._addField(fields, attrname="DADDRESS1", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="DADDRESS2", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="DADDRESS3", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="DADDRESS4", fieldtype="string", WIDTH="60")
            self._addField(fields, attrname="DATTENTION", fieldtype="string", WIDTH="70")
            self._addField(fields, attrname="DPHONE1", fieldtype="string", WIDTH="30")
            self._addField(fields, attrname="DFAX1", fieldtype="string", WIDTH="30")
            self._addField(fields, attrname="ATTACHMENTS", fieldtype="bin.hex", SUBTYPE="Binary", WIDTH="8")
            self._addField(fields, attrname="NOTE", fieldtype="bin.hex", SUBTYPE="Binary", WIDTH="8")
            self._addField(fields, attrname="TRANSFERABLE", fieldtype="string", WIDTH="1")
            self._addField(fields, attrname="UPDATECOUNT", fieldtype="i4")
            self._addField(fields, attrname="PRINTCOUNT", fieldtype="i4")
            self._addField(fields, attrname="DOCNOSETKEY", fieldtype="i8", required="true")
            self._addField(fields, attrname="NEXTDOCNO", fieldtype="string", WIDTH="20")
            self._addField(fields, attrname="CHANGED", fieldtype="string", required="true", WIDTH="1")

            sdsDocDetail = self._addField(fields, attrname="sdsDocDetail", fieldtype="nested")
            sdsDocDetail_fields = ET.SubElement(sdsDocDetail, "FIELDS")
            self._addField(sdsDocDetail_fields, attrname="DTLKEY", fieldtype="i4", required="true")
            self._addField(sdsDocDetail_fields, attrname="DOCKEY", fieldtype="i4", required="true")
            self._addField(sdsDocDetail_fields, attrname="SEQ", fieldtype="i4")
            self._addField(sdsDocDetail_fields, attrname="STYLEID", fieldtype="string", WIDTH="5")
            self._addField(sdsDocDetail_fields, attrname="NUMBER", fieldtype="string", WIDTH="5")
            self._addField(sdsDocDetail_fields, attrname="ITEMCODE", fieldtype="string", WIDTH="30")
            self._addField(sdsDocDetail_fields, attrname="LOCATION", fieldtype="string", WIDTH="20")
            self._addField(sdsDocDetail_fields, attrname="BATCH", fieldtype="string", WIDTH="30")
            self._addField(sdsDocDetail_fields, attrname="PROJECT", fieldtype="string", WIDTH="20")
            self._addField(sdsDocDetail_fields, attrname="DESCRIPTION", fieldtype="string", WIDTH="200")
            self._addField(sdsDocDetail_fields, attrname="DESCRIPTION2", fieldtype="string", WIDTH="200")
            self._addField(sdsDocDetail_fields, attrname="DESCRIPTION3", fieldtype="bin.hex", SUBTYPE="Binary", WIDTH="8")
            # self._addField(sdsDocDetail_fields, attrname="PERMITNO", fieldtype="string", WIDTH="20")
            self._addField(sdsDocDetail_fields, attrname="RECEIVEQTY", fieldtype="fixedFMT", DECIMALS="4", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="RETURNQTY", fieldtype="fixedFMT", DECIMALS="4", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="QTY", fieldtype="fixedFMT", DECIMALS="4", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="UOM", fieldtype="string", WIDTH="10")
            self._addField(sdsDocDetail_fields, attrname="RATE", fieldtype="fixedFMT", DECIMALS="4", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="SQTY", fieldtype="fixedFMT", DECIMALS="4", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="SUOMQTY", fieldtype="fixedFMT", DECIMALS="4", WIDTH="19")
            # self._addField(sdsDocDetail_fields, attrname="OFFSETQTY", fieldtype="fixedFMT", DECIMALS="4", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="UNITPRICE", fieldtype="fixedFMT", DECIMALS="8", WIDTH="19")
            # self._addField(sdsDocDetail_fields, attrname="DELIVERYDATE", fieldtype="date")
            self._addField(sdsDocDetail_fields, attrname="DISC", fieldtype="string", WIDTH="20")
            self._addField(sdsDocDetail_fields, attrname="TAX", fieldtype="string", WIDTH="10")
            self._addField(sdsDocDetail_fields, attrname="TAXRATE", fieldtype="string", WIDTH="20")
            self._addField(sdsDocDetail_fields, attrname="TAXAMT", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="LOCALTAXAMT", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="TAXINCLUSIVE", fieldtype="i2")
            self._addField(sdsDocDetail_fields, attrname="AMOUNT", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="LOCALAMOUNT", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="PRINTABLE", fieldtype="string", WIDTH="1")
            self._addField(sdsDocDetail_fields, attrname="FROMDOCTYPE", fieldtype="string", WIDTH="2")
            self._addField(sdsDocDetail_fields, attrname="FROMDOCKEY", fieldtype="i4")
            self._addField(sdsDocDetail_fields, attrname="FROMDTLKEY", fieldtype="i4")
            self._addField(sdsDocDetail_fields, attrname="TRANSFERABLE", fieldtype="string", WIDTH="1")
            self._addField(sdsDocDetail_fields, attrname="REMARK1", fieldtype="string", WIDTH="200")
            self._addField(sdsDocDetail_fields, attrname="REMARK2", fieldtype="string", WIDTH="200")
            self._addField(sdsDocDetail_fields, attrname="INITIALPURCHASECOST", fieldtype="fixedFMT", DECIMALS="2", WIDTH="19")
            self._addField(sdsDocDetail_fields, attrname="CHANGED", fieldtype="string", required="true", WIDTH="1")

            sdsSerialNumber = self._addField(sdsDocDetail_fields, attrname="sdsSerialNumber", fieldtype="nested")
            sdsSerialNumber_fields = ET.SubElement(sdsSerialNumber, "FIELDS")
            self._addField(sdsSerialNumber_fields, attrname="SERIALNUMBER", fieldtype="string", required="true", WIDTH="30")
            sdsSerialNumber_params = ET.SubElement(sdsSerialNumber, "PARAMS")
            sdsSerialNumber_params.set("MD_SEMANTICS", "2")
            sdsSerialNumber_params.set("LCID", "0")

            params = ET.SubElement(metadata, "PARAMS")
            params.set("MD_SEMANTICS", "2")
            params.set("LCID", "0")
            param = ET.SubElement(params, "PARAM")
            param.set("Name", "MD_FIELDLINKS")
            param.set("Value", "52 1 2")
            param.set("Type", "IntArray")

        
            rowdata = ET.SubElement(datapacket, "ROWDATA")
        
            # for rec in recs:
            po_code = rec['name']
            supplier_code = str(rec['partner_id'].id)
            # parsedate = ""
            if isBulk:
                zipfile_date  = datetime.now().strftime("%Y%m%d")
            else:
                zipfile_date = (datetime.strptime(rec['confirm_date'], "%Y-%m-%d %H:%M:%S")).strftime("%Y%m%d")
            parsedate = (datetime.strptime(rec['confirm_date'], "%Y-%m-%d %H:%M:%S")).strftime("%Y%m%d")
            row = ET.SubElement(rowdata, "ROW")
            row.set("DOCKEY", "")
            s = list(rec['name'])
            s[0] = 'D'
            docno = "".join(s)
            row.set("DOCNO", docno)
            row.set("DOCNOEX", "")
            row.set("DOCDATE", parsedate)
            row.set("POSTDATE", parsedate)
            row.set("TAXDATE", parsedate)

            lines = self.env['purchase.order.line'].search([('order_id', '=', rec['id'])])
            buyer = self.env['res.partner'].search([('id', '=', lines[0].buyer_id)])
            if buyer:
                code = str(buyer.id)
                if buyer.ref:
                    code = buyer.ref
                row.set("CODE", code)
                row.set("COMPANYNAME", str(buyer.name or ''))
                row.set("ADDRESS1", str(buyer.street or ''))
                row.set("ADDRESS2", str(buyer.city or ''))
            else:
                code = str(rec['partner_id'].id)
                if rec['partner_id'].ref:
                    code = rec['partner_id'].ref
                row.set("CODE", code)
                row.set("COMPANYNAME", str(rec['partner_id'].name or ''))
                row.set("ADDRESS1", str(rec['partner_id'].street or ''))
                row.set("ADDRESS2", str(rec['partner_id'].city or ''))
            
            row.set("ADDRESS3", "")
            row.set("ADDRESS4", "")
            row.set("PHONE1", "")
            row.set("FAX1", "")
            row.set("ATTENTION", "----")
            row.set("AREA", "----")
            row.set("AGENT", "")
            row.set("PROJECT", "----")
            row.set("TERMS", "")
            row.set("CURRENCYCODE", "----")
            row.set("CURRENCYRATE", "1.0000")
            row.set("SHIPPER", "----")
            row.set("DESCRIPTION", "Delivery Order")
            row.set("COUNTRY", "")
            row.set("CANCELLED", "F")
            row.set("DOCAMT", "0.00")
            row.set("LOCALDOCAMT", "0.00")
            # row.set("D_DOCNO", "")
            # row.set("D_PAYMENTMETHOD", "")
            # row.set("D_CHEQUENUMBER", "")
            # row.set("D_PAYMENTPROJECT", "")
            # row.set("D_BANKCHARGE", "")
            row.set("D_AMOUNT", "0.00")
            row.set("VALIDITY", "")
            row.set("DELIVERYTERM", "")
            row.set("CC", "")
            row.set("DOCREF1", str(rec['pda_do_id']))
            row.set("DOCREF2", "")
            row.set("DOCREF3", "")
            row.set("DOCREF4", "")
            row.set("BRANCHNAME", "")
            row.set("DADDRESS1", "")
            row.set("DADDRESS2", "")
            row.set("DADDRESS3", "")
            row.set("DADDRESS4", "")
            row.set("DATTENTION", "")
            row.set("DPHONE1", "")
            row.set("DFAX1", "")
            row.set("ATTACHMENTS", "")
            row.set("NOTE", "")
            row.set("TRANSFERABLE", "T")
            row.set("UPDATECOUNT", "")
            row.set("PRINTCOUNT", "0")
            row.set("DOCNOSETKEY", "0")
            row.set("NEXTDOCNO", "0")
            row.set("CHANGED", "F")

            sdsdocdetail = ET.SubElement(row, "sdsDocDetail")
            seq=1

            itemcode = ""
            location = ""
            description = ""
            qty = 0.0
            uom = ""
            unitprice = 0.0
            amount = ""
            localamount = ""
            for line in lines:
                if seq == 1:
                    itemcode = str(line['product_id'].id)
                    if line['product_id'].default_code:
                        itemcode = line['product_id'].default_code
                    location = str(rec['partner_id'].name or '----')
                    description = str(line['product_id'].name_template)
                    uom = str(line['product_uom'].name)
                unitprice += line['price_unit']
                qty += line['crate_net']
                seq+=1
            suomqty = rec['total_male'] + rec['total_female'] + rec['birdcrate_mix_total'] + rec['birdcrate_b_total']

            rows = ET.SubElement(sdsdocdetail, "ROWsdsDocDetail")
            rows.set("DTLKEY", "")
            rows.set("DOCKEY", "")
            rows.set("SEQ", "1")
            rows.set("STYLEID", "")
            rows.set("NUMBER", "")
            rows.set("ITEMCODE", itemcode)
            rows.set("LOCATION", location)
            rows.set("BATCH", "")
            rows.set("PROJECT", "----")
            rows.set("DESCRIPTION", description)
            rows.set("DESCRIPTION2", "")
            rows.set("DESCRIPTION3", "")
            rows.set("RECEIVEQTY", "")
            rows.set("RETURNQTY", "")
            rows.set("QTY", str(qty))
            rows.set("UOM", uom.upper())
            rows.set("RATE", "")
            rows.set("SQTY", "")
            rows.set("SUOMQTY", str(int(suomqty)))
            # rows.set("UNITPRICE", str(unitprice) or '0.00')
            rows.set("UNITPRICE", "0.0")
            rows.set("DISC", "")
            rows.set("TAX", "")
            rows.set("TAXRATE", "")
            rows.set("TAXAMT", "0.00")
            rows.set("LOCALTAXAMT", "0.00")
            rows.set("TAXINCLUSIVE", "0")
            # rows.set("AMOUNT", "{0:,.2f}".format(round(unitprice,2)))
            rows.set("AMOUNT", "0.0")
            # rows.set("LOCALAMOUNT", "{0:,.2f}".format(round(unitprice,2)))
            rows.set("LOCALAMOUNT", "0.0")
            rows.set("PRINTABLE", "T")
            rows.set("FROMDOCTYPE", "")
            rows.set("FROMDOCKEY", "")
            rows.set("FROMDTLKEY", "")
            rows.set("TRANSFERABLE", "T")
            rows.set("REMARK1", "")
            rows.set("REMARK2", "")
            rows.set("INITIALPURCHASECOST", "")
            rows.set("CHANGED", "F")

            datas = ET.tostring(datapacket)
            file = filename + '.' + po_code + '.' + supplier_code + '.xml'
            files_pack.append(file)
            f =  open(files_path + "/" + file, "wb")
            f.write(datas)
            f.close()

        zip_file = filename + "_" + zipfile_date + ".zip"

        # writing files to a zipfile 
        with ZipFile(files_path + "/" + zip_file, 'w') as newzip:
            for file in files_pack:
                name = files_path + "/" + file
                newzip.write(name, basename(name))
            newzip.close()

        
        # print etree.tostring(root)
        soupese = self.env['soupese.common_wizard']
        wizard_rec = soupese.create({'file_name': zip_file, 
                                     'file': base64.b64encode(open(files_path + "/" + zip_file,'rb').read())})

        view_id = self.env['ir.model.data'].get_object_reference('soupese_base', 'post_export_view')[1]
        return {
            'name':'Export to XML',
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'soupese.common_wizard',
            'view_id': view_id,
            'res_id': wizard_rec.id,
            'target': 'new',
            'context': {},
            }
