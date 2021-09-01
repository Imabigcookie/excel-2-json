import Vue from 'vue'
import VueCompositionAPI, { createApp } from '@vue/composition-api'
import App from './App.vue'
import './styles/app.css'

Vue.config.productionTip = false

Vue.use(VueCompositionAPI)

const app = createApp({
  render: (h) => h(App),
})

app.mount('#app')
