import Vue from 'vue'
import Vuex from 'vuex'

Vue.use(Vuex)

const store = new Vuex.Store({
  state: {
    currency_data: []
  },
  mutations: {
    setCurrencyData (state, data) {
      state.currency_data = data
    },
    destroyCurrencyData (state) {
      state.currency_data = []
    }
  },
  actions: {
    setCurrencyData (context, data) {
      context.commit('setCurrencyData', data)
    },
    destroyCurrencyData (context) {
      context.commit('destroyCurrencyData')
    }
  }
})

export default store
