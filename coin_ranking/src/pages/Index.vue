<template>
  <div class="q-pa-lg">
    <q-table
      title="Top Crypto Currency List"
      :data="currency_details"
      :columns="columns"
      :pagination.sync="pagination"
      row-key="name">
      <template v-slot:body="props">

        <q-tr :props="props">
          <q-td key="iconUrl" :props="props" style="text-align: center">
            <q-icon class="text-dark">
              <img :src="props.row.iconUrl" style="height: 25px;" alt=""/>
            </q-icon>
            <!--              <q-icon :name="img:props.row.iconUrl"></q-icon>-->
          </q-td>
          <q-td key="name" :props="props" style="text-align: center">
            {{props.row.name}}
          </q-td>
          <q-td key="symbol" :props="props" style="text-align: center">
            {{props.row.symbol}}
          </q-td>
          <q-td key="price" :props="props" style="text-align: center">
            ${{props.row.price}}
          </q-td>
          <q-td key="change" :props="props" style="text-align: center">
            {{props.row.change}}%
          </q-td>
        </q-tr>
      </template>
    </q-table>
    <!--    <q-card class="my-card">-->
    <!--      <q-card-section>-->
    <!--        <div class="text-h6">Top Crypto Currency List</div>-->
    <!--      </q-card-section>-->

    <!--       <div class="q-pa-md">-->

    <!--  </div>-->
    <!--    </q-card>-->
  </div>
</template>
<script>
import axios from 'axios'

export default {
  name: 'PageIndex',
  data () {
    return {
      polling: null,
      pagination: {
        rowsPerPage: 10 // current rows per page being displayed
      },
      columns: [
        {
          name: 'iconUrl',
          required: true,
          label: 'Currency',
          align: 'center',
          style: { 'font-size': '15px' },
          field: row => row.iconUrl,
          sortable: true
        },
        {
          name: 'name',
          align: 'center',
          label: 'Currency Name',
          field: 'name',
          sortable: true
        },
        {
          name: 'symbol',
          align: 'center',
          label: 'Currency Symbol',
          field: 'symbol',
          sortable: true
        },
        {
          name: 'price',
          align: 'center',
          label: 'Currency Price',
          field: 'price',
          sortable: true
        },
        {
          name: 'change',
          align: 'center',
          label: 'Currency Price Change',
          field: 'change',
          sortable: true
        }
      ],
      currency_details: []
    }
  },
  methods: {
    getCurrencyDetails () {
      axios.get('https://api.coinranking.com/v1/public/coins/?limit=100')
        .then(response => {
          // debugger
          this.currency_details = response.data.data.coins
          this.$store.dispatch('setCurrencyData', response.data.data.coins)
          this.pollData()
        })
        .catch(e => {
          this.errors.push(e)
        })
    },
    pollData () {
      this.$store.dispatch('destroyCurrencyData')
      this.polling = setInterval(() => {
        this.getCurrencyDetails()
      }, 30 * 1000)
    }

  },
  mounted: function () {
    this.getCurrencyDetails()
  }

}
</script>
<style lang="stylus">
  .my-card {
    width: 100%
  }

  /*max-width: 500px*/
</style>
