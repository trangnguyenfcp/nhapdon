<template>
  <v-container>
    <v-row class="text-center">
      <v-col cols="12">
        <b>Tính công</b>
      </v-col>
    </v-row>
    <v-row>
      <v-form>
        <v-row class="ma-5">
          
          <v-col  class="col-4 col-lg-4 col-md-4 col-sm-4">
            <v-label>
              Sản phẩm phụ 1
            </v-label>
            <v-text-field placeholder="Có hoặc không" v-model="jsonData"></v-text-field>
          </v-col>
        </v-row>
      </v-form>
    </v-row>
    <v-row class="ml-3 mr-3">
      <v-col cols="12">
        <v-row>
          <v-col>
            <v-btn
            @click="write()"
            >
              Tính
            </v-btn>
          </v-col>
          <v-col>
            <v-label>
              Tình trang: 
            </v-label>
            <span>
              {{status}}
            </span>
          </v-col>
        </v-row>
      </v-col>
    </v-row>
  </v-container>
</template>

<script>
const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('@/client_secret.json');
  export default {
    name: 'HelloWorld',

    data: () => ({
      jsonData: "",
      status: 'chưa',
      listData: [],
      sheet: {},
      headers: [],
      shortenedNameSheet: {},
      fromRow: 500,
    }),
    methods:{
      async checkAndWriteIfNotExists(columnIndex, rowIndex, data) {
        try {
          
          // Load all cells in the specified column
          await this.sheet.loadCells({
            startRowIndex: 0,
            endRowIndex: this.sheet.rowCount,
            startColumnIndex: columnIndex,
            endColumnIndex: columnIndex + 1,
          });

          let cellWithValue = false;

          // Iterate through each cell in the column
          for (let i = rowIndex; i < this.sheet.rowCount; i++) {
            const cell = this.sheet.getCell(i, columnIndex);

            // Check if the cell has a value
            if (cell.value !== null && cell.value !== '') {
              cellWithValue = true;

              // Check if the cell contains the desired value
              if (cell.value === data) {
                console.log(`Value '${data}' already exists in the column`);
                return;
              }
            } else {
              // If an empty cell is found, write data to it
              cell.value = data;
              await this.sheet.saveUpdatedCells();
              console.log(`Data '${data}' written to cell (${i}, ${columnIndex})`);
              return;
            }
          }

          // If all cells in the column have values and none contain the desired value, write data to the next empty cell
          if (!cellWithValue) {
            const lastRowIndex = this.sheet.rowCount;
            const lastEmptyCell = this.sheet.getCell(lastRowIndex, columnIndex);
            lastEmptyCell.value = data;
            await this.sheet.saveUpdatedCells();
            console.log(`Data '${data}' written to next empty cell in column ${columnIndex}`);
          }
        } catch (error) {
          console.error('Error:', error);
        }
      },

      async checkAndWriteData(column1Index, column2Index, rowIndex, orderNumber, value) {
        try {
          // Load cells of both columns
          await this.sheet.loadCells({
            startRowIndex: 0,
            endRowIndex: this.sheet.rowCount,
            startColumnIndex: 0,
            endColumnIndex: 20,
          });

          let modified = false;
          // Iterate through each row
          for (let i = rowIndex; i < this.sheet.rowCount; i++) {
            const cell1 = this.sheet.getCell(i, column1Index);
            const cell2 = this.sheet.getCell(i, column2Index);

            // Check if the value of cell2 equals orderNumber
            if (cell2.value === orderNumber) {
              // Write value to cell1 if condition is met
              cell1.value = value;
              modified = true;
              console.log(`Data '${value}' written to cell (${i}, ${column1Index})`);
            }
          }

          if (!modified) {
            console.log(`No cells found where value in column ${column2Index} equals '${orderNumber}'`);
          } else {
            // Save changes to the spreadsheet
            await this.sheet.saveUpdatedCells();
          }
        } catch (error) {
          console.error('Error:', error);
        }
      },

      async checkAndReturnValue(column1Index, column2Index, data2) {
        try {
          // Load cells of both columns
          await this.shortenedNameSheet.loadCells({
            startRowIndex: 0,
            endRowIndex: 200,
            startColumnIndex: 0,
            endColumnIndex: 10,
          });

          // Iterate through each row
          for (let i = 0; i < 200; i++) {
            const cell1 = this.shortenedNameSheet.getCell(i, column1Index);
            const cell2 = this.shortenedNameSheet.getCell(i, column2Index);
            

            // Check if the value of cell2 equals data2
            if (cell2.value === data2) {
              // Return the value of cell1 if condition is met
              return cell1.value;
            }
          }

          // Return null if no matching value is found
          return null;
        } catch (error) {
          console.error('Error:', error);
          return null;
        }
      },

      async write(){
        const data = JSON.parse(this.jsonData).data.data;
        const listOrder = JSON.parse(this.jsonData).data.hierarchy.structure;
        console.log(listOrder)
        const listKey = Object.keys(data);
        const listOrderKey = Object.keys(listOrder);
        

        for (const key of listOrderKey) {
          if (key.startsWith('order_')) {
            const columnIndex = this.headers.findIndex(header => header === "Mã đơn hàng");
            const orderId = key.split('_')[1];
            await this.checkAndWriteIfNotExists(columnIndex, this.fromRow , orderId);
          }
        }

        
        for (const key of listKey) {
          if (key.startsWith('orderOperation_OrderOperation_ORDERLOGIC_')) {
            const listProduct = data[key].fields.orderLineIds;
            const orderId = data[key].fields.orderId;
            let name = "";
            // let sku = "";
            let quantity = 0;
            let price = 0;
            let status = "";
            let check = false;
            for (const productId of listProduct) {
              let date = data["orderItem_" + productId].fields.delivery?.desc;
              if (date && date.startsWith("Đã giao ngày")) {
                date = date.replace(/Đã giao ngày /g, '');
                date = date.replace(/ thg /g, '-');
                date = date.replace(/ /g,'-');
                const y = date.split('-')[2];
                const m = date.split('-')[1];
                const d = date.split('-')[0];
                date = y + '-' + m + '-' + d
                check = new Date(date) < new Date('2024-02-29');
              }
              if (data["orderItem_" + productId].fields.status === "Đã hủy đơn") check = true;
              if (check) continue;
              const checkName = await this.checkAndReturnValue(1, 0, data["orderItem_" + productId].fields.title);
              if (checkName) data["orderItem_" + productId].fields.title = checkName;
              if (data["orderItem_" + productId].fields.title == "Mobi 10k" || data["orderItem_" + productId].fields.title == "Vina 10k") check = true;
              if (name != "") name += " + ";
              name += data["orderItem_" + productId].fields.title;

              // if (data["orderItem_" + productId].fields.sku.skuText == "") data["orderItem_" + productId].fields.sku.skuText = "0" ;
              // if (sku != "") sku += " + " ;
              // sku += data["orderItem_" + productId].fields.sku.skuText;
              
              if (quantity != "") quantity += " + " ;
              quantity += data["orderItem_" + productId].fields.quantity;

              let itemPrice = data["orderItem_" + productId].fields.price;
              itemPrice = itemPrice.replace(/\./g,'');
              itemPrice = itemPrice.replace(['₫ '],'');
              price += Math.round(Number(itemPrice)/1000)
              if (data["orderItem_" + productId].fields.status === "Thủ tục hoàn tiền đã hoàn tất.") status = "Hoàn tiền"
              
            }
            if (check) continue;
            const columnIndex1 = this.headers.findIndex(header => header === "Mã đơn hàng");
            const columnIndex2 = this.headers.findIndex(header => header === "Tên sản phẩm");
            // const columnIndex3 = this.headers.findIndex(header => header === "Phân loại");
            const columnIndex4 = this.headers.findIndex(header => header === "SL");
            const columnIndex5 = this.headers.findIndex(header => header === "Giá SP");
            const columnIndex6 = this.headers.findIndex(header => header === "Đơn vị vận chuyển");
            await this.checkAndWriteData(columnIndex2, columnIndex1, this.fromRow , orderId, name);
            // await this.checkAndWriteData(columnIndex3, columnIndex1, this.fromRow , orderId, sku);
            await this.checkAndWriteData(columnIndex4, columnIndex1, this.fromRow , orderId, quantity);
            await this.checkAndWriteData(columnIndex5, columnIndex1, this.fromRow , orderId, price);
            await this.checkAndWriteData(columnIndex6, columnIndex1, this.fromRow , orderId, status);
          }
        }

        for (const key of listKey) {
          if (key.startsWith('orderExtraSummary_OrderExtraSummary_ORDERLOGIC_')) {
            const orderId = key.split('_')[3];
            let totalPrice = data[key].fields.value;
            totalPrice = totalPrice.replace(/\./g,'');
            totalPrice = totalPrice.replace(['₫ '],'');
            const columnIndex1 = this.headers.findIndex(header => header === "Mã đơn hàng");
            const columnIndex2 = this.headers.findIndex(header => header === "Trả trước");
            await this.checkAndWriteData(columnIndex2, columnIndex1, this.fromRow, orderId, Math.round(Number(totalPrice)/1000));
          }
        }
        
      },
			async accessSpreadSheet() {
				const doc = new GoogleSpreadsheet('1caITNjs8pUec8gLQLaRBgupzkCvDrLGMCP_fmNRfGss');
				await doc.useServiceAccountAuth(creds);
				await doc.loadInfo(); 
				this.sheet = doc.sheetsByTitle['Tháng 03-2024'];
        await this.sheet.loadHeaderRow();
        this.headers = this.sheet.headerValues;
        this.shortenedNameSheet = doc.sheetsByTitle['Tên sản phẩm rút gọn'];
			}
		},
    created() {
			this.accessSpreadSheet();
		},
    
  }
</script>
<style>
  .v-form {
    width: 100%;
  }
</style>
