<template>
<div class="container">
  <h1>Excel to JSON</h1>
  <input
    class="form-control mt-5"
    ref="input"
    type="file"
    multiple
    accept=".xlsx, .xls"
    @input="onInput"
  >

  <div class="errors-block mt-4">
    <div
      class="text-danger mt-2"
      v-for="(error, index) in errors"
      :key="index"
    >{{ error }}</div>
  </div>

  <button
    class="btn btn-success mt-4"
    @click="clearErrors"
  >
    Очистить ошибки
  </button>
</div>
</template>

<script>
import downloadFile from 'js-file-download'

export default {
  name: 'App',
  data: () => ({
    errors: [],
    files: [],
  }),
  methods: {
    clearErrors() {
      this.errors = []
      this.$refs.input.files = []
    },

    async onInput() {
      this.files = [ ...this.$refs.input.files ]
      this.errors = []

      const promises = this.files.map(file => this.converteFile(file))

      await Promise.all(promises)
    },

    calculateSum(range) {
      return range.reduce((acc, row, i) => {
        if (i === 0) return acc

        return acc + row[8]
      }, 0)
    },

    async converteFile(file) {
      try {
        const openedFile = await XlsxPopulate.fromDataAsync(file)
        const workbook = openedFile.sheet(0)
        const range = workbook.usedRange().value()
        const totalSum =this.calculateSum(range)
        const result = {}
        const last = {
          name: '',
          value: 0
        }

        range.forEach((row, i) => {
          if (i === 0) return

          const resultName = `${i} ${row[2]}`
          const resultValue = +((row[8] / totalSum) * 100).toFixed(3)

          result[resultName] = resultValue
          last.name = resultName
          last.value = resultValue
        });

        const sumWithoutLast = this.calculateSum(range.slice(0, -1))
        const lastValue = +((1 - sumWithoutLast / totalSum) * 100).toFixed(3)

        if (Math.abs(lastValue - last.value) > 0.3) {
          this.errors.push(`${file.name}: JSON не сформирвоан, превышена дельта последней позиции`)

          return
        }

        result[last.name] = lastValue

        downloadFile(JSON.stringify(result, null, '\n'), `${file.name.split('.')[0]}.json`);
      } catch (error) {
        console.error(error)
        this.errors.push(`${file.name}: Ошибка программы`)
      } finally {
      }
    }
  },
}
</script>
