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
              Sản phẩm chính
            </v-label>
            <v-text-field placeholder="Sản phẩm chính" :error-messages="rules.required" required type="number" v-model="price.mainProduct"></v-text-field>
          </v-col>
          <v-col  class="col-4 col-lg-4 col-md-4 col-sm-4">
            <v-label>
              Sản phẩm phụ 1
            </v-label>
            <v-text-field placeholder="Có hoặc không" type="number" v-model="price.byProduct1"></v-text-field>
          </v-col>
          <v-col  class="col-4 col-lg-4 col-md-4 col-sm-4">
            <v-label>
              Sản phẩm phụ 2
            </v-label>
            <v-text-field placeholder="Có hoặc không" type="number" v-model="price.byProduct2"></v-text-field>
          </v-col>
          <v-col  class="col-4">
            <v-label>
            Mã sàn  
            </v-label>
            <v-text-field placeholder="Có hoặc không" type="number" v-model="discount.lzdDiscount"></v-text-field>
          </v-col>
          <v-col  class="col-4">
            <v-label>
            Mã bank  
            </v-label>
            <v-text-field placeholder="Có hoặc không" type="number" v-model="discount.bankDiscount"></v-text-field>
          </v-col>
          <v-col  class="col-4">
            <v-label>
            Gạt hoàn tiền  
            </v-label>
            <v-text-field placeholder="Có hoặc không" type="number" v-model="discount.refundDiscount" :error-messages="rules.refundMax"></v-text-field>
          </v-col>
          <v-col  class="col-4">
            <v-label>
            Voucher tích lũy  
            </v-label>
            <v-text-field placeholder="Có hoặc không" type="number" v-model="discount.cumulativeDiscount"></v-text-field>
          </v-col>
          <v-col  class="col-4">
            <v-label>
            Voucher freeship  
            </v-label>
            <v-select placeholder="Có hoặc không" :items="freeshipDiscounts" v-model="discount.freeshipDiscount"></v-select>
          </v-col>
          <v-col  class="col-4">
            <v-label>
            Ngày đặt đơn
            </v-label>
            <v-menu
              ref="menu"
              v-model="menu"
              :close-on-content-click="false"
              :return-value.sync="date"
              transition="scale-transition"
              offset-y
              min-width="290px"
            >
              <template v-slot:activator="{ on, attrs }">
                <v-text-field
                  v-model="date"
                  readonly
                  v-bind="attrs"
                  v-on="on"
                ></v-text-field>
              </template>
              <v-date-picker v-model="date" no-title scrollable>
                <v-spacer></v-spacer>
                <v-btn text color="primary" @click="menu = false">Cancel</v-btn>
                <v-btn text color="primary" @click="$refs.menu.save(date)">OK</v-btn>
              </v-date-picker>
            </v-menu>
          </v-col>
          <!-- <v-col  class="col-6">
            <v-label>
              Tổng đơn:
            </v-label>
            <span>
              {{price.totalPurchasePrice}}
            </span>
          </v-col>
          <v-col  v-if="price.byProduct1 >= 200 || price.byProduct2 >= 200" class="col-6">
            <span>
              Trừ 20.000 tiền công
            </span>
          </v-col>
          <v-col  class="col-6">
          </v-col> -->
        </v-row>
      </v-form>
    </v-row>
    <v-row class="ml-3 mr-3">
      <v-col cols="12">
        <v-row>
          <v-col>
            <v-btn
            @click="calculate()"
            >
              Tính
            </v-btn>
          </v-col>
          <v-col>
            <v-label>
              Công: 
            </v-label>
            <span>
              {{wages}}
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
      rules : {
        required: "Bắt buộc nhập",
        refundMax: "<= 100",
      },
      provisionalWagesRows: {},
      priceLevelRows: {},
      extraWagesRows: {},
      message: '',
      price: {
        mainProduct: '',
        byProduct1: '',
        byProduct2: '',
        totalPurchasePrice: 0,

      },
      discount: {
        lzdDiscount: '',
        bankDiscount: '',
        freeshipDiscount: '',
        cumulativeDiscount: '',
        refundDiscount: '',
        provisionalDiscount: '',
        discountCal: '',
      },
      freeshipDiscounts: [
        {value: 0, text: "0"},
        {value: 15, text: "15"},
        {value: 20, text: "20"}
      ],
      extraWages: 0,
      provisionalWages: 0,
      wages: '',
      date: new Date().toISOString().substr(0, 10),
      menu: false,

    }),
    methods:{
      clearData() {
        this.message = "";
        this.extraWages = 0;
        this.provisionalWages= 0;
        this.wages = '';
      },
      calculateProvisionalWages() {
        this.clearData();
        const lzdDiscount = this.discount.lzdDiscount > 0 ? parseInt(this.discount.lzdDiscount) : 0;
        const bankDiscount = this.discount.bankDiscount > 0 ? parseInt(this.discount.bankDiscount) : 0;
        this.discount.provisionalDiscount = lzdDiscount + bankDiscount;
        this.calculateTotalPurchasePrice();
        let priceLevel = '';
        for (let i =0; i < this.priceLevelRows.length; i++){
          const lowestPrice = parseInt(this.priceLevelRows[i].lowestPrice);
          const hightestPrice = parseInt(this.priceLevelRows[i].hightestPrice);
          if ((this.price.totalPurchasePrice > lowestPrice) && (this.price.totalPurchasePrice <= hightestPrice)) {
            priceLevel = this.priceLevelRows[i].priceLevel;
            break;
          }
        }
        if (!priceLevel) {
          this.message = "Tổng đơn < 500, > 5900 liên hệ tele Ken";
          return;
        }
        if (this.discount.refundDiscount > 100) {
          this.message = "Gạt hoàn tiền <= 100";
          return;
        }
        if (this.discount.provisionalDiscount < 100 || this.discount.provisionalDiscount > 1500) {
          this.message = "Mã sàn + mã bank < 100, > 1500 liên hệ tele Ken";
          return;
        }
        for (let i =0; i < this.provisionalWagesRows.length - 1; i++){
          if (this.provisionalWagesRows[i].priceLevel === priceLevel){
            if (this.discount.provisionalDiscount >= this.provisionalWagesRows[i].discount && this.discount.provisionalDiscount < this.provisionalWagesRows[i + 1].discount) {
              this.discount.discountCal = this.provisionalWagesRows[i].discount;
            }
          }
        }
        for (let i = 0; i < this.provisionalWagesRows.length; i++) {
          if ((this.provisionalWagesRows[i].priceLevel == priceLevel) && (this.provisionalWagesRows[i].discount ==  this.discount.discountCal)) {
            this.provisionalWages = this.provisionalWagesRows[i].provisionalWages;
          }
        }

        for (let i = 0; i < this.extraWagesRows.length; i++) {
          const dateFrom = new Date(this.extraWagesRows[i].dateFrom);
          const dateTo = new Date(this.extraWagesRows[i].dateTo);
          const dateChosen = new Date(this.date);
          if (dateChosen >= dateFrom && dateChosen <= dateTo && this.extraWagesRows[i].priceLevel == priceLevel) {
            this.extraWages = this.extraWagesRows[i].extraWages;
          }
        }
        if (this.extraWages) return;
        for (let i = 0; i < this.extraWagesRows.length; i++) {
          if (!this.extraWagesRows[i].dateFrom && !this.extraWagesRows[i].dateTo && this.extraWagesRows[i].priceLevel == priceLevel) {
            this.extraWages = this.extraWagesRows[i].extraWages;
          }
        }
      },
      calculateTotalPurchasePrice(){
        const mainProduct = this.price.mainProduct > 0 ? parseInt(this.price.mainProduct) : 0;
        const byProduct1 = this.price.byProduct1 > 0 ? parseInt(this.price.byProduct1) : 0;
        const byProduct2 = this.price.byProduct2 > 0 ? parseInt(this.price.byProduct2) : 0;
        this.price.totalPurchasePrice =  mainProduct + byProduct1 + byProduct2;
      },
      calculate(){
        this.calculateProvisionalWages();
        if (this.message !== "") {
          this.wages = this.message;
          return;
        }
        
        const provisionalWages = this.provisionalWages ? parseInt(this.provisionalWages) : 0;
        const cumulative = this.discount.cumulativeDiscount ? parseInt(this.discount.cumulativeDiscount*40/100) : 0;
        const freeshipDiscount = this.discount.freeshipDiscount ? parseInt(this.discount.freeshipDiscount) : 0;
        const refundDiscount = this.discount.refundDiscount ? parseInt(this.discount.refundDiscount) : 0;
        const fee = (this.price.byProduct1 > 200 || this.price.byProduct2 > 200) ? 20 : 0;
        const remainWages = this.discount.provisionalDiscount - this.discount.discountCal;
        const extraWages = parseInt(this.extraWages);
        this.wages = provisionalWages + freeshipDiscount + refundDiscount + cumulative - fee + remainWages + extraWages;
      },
			async accessSpreadSheet() {
				const doc = new GoogleSpreadsheet('106OJmbNkC6Fp-ZOD5dMzXvz180Nr_jzcM87l235utfc');
				await doc.useServiceAccountAuth(creds);
				await doc.loadInfo(); 
				const sprovisionalWagesSheet = doc.sheetsByIndex[1];
        const priceLevelSheet = doc.sheetsByIndex[2];
        const extraWagesSheet = doc.sheetsByIndex[3];
				const provisionalWagesRows = await sprovisionalWagesSheet.getRows({
					offset: 0
				});
        const priceLevelRows = await priceLevelSheet.getRows({
					offset: 0
				});
        const extraWagesRows = await extraWagesSheet.getRows({
					offset: 0
				});
        this.provisionalWagesRows = provisionalWagesRows;
        this.priceLevelRows = priceLevelRows;
        this.extraWagesRows = extraWagesRows;        
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
