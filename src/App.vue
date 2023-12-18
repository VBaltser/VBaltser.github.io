<template>
    <BarCodeGenerator msg="Welcome to Your Vue.js + TypeScript App" />
</template>

<script lang="ts">
import { defineComponent } from "vue";
import BarCodeGenerator from "./components/BarcodeGenerator.vue";

export default defineComponent({
    name: "App",
    components: {
        BarCodeGenerator,
    },
});
</script>

<style>
#app {
    font-family: Avenir, Helvetica, Arial, sans-serif;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
    margin-top: 60px;
}
</style>
