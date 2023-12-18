<template>
    <div class="label-container" />
</template>

<script lang="ts">
import { defineComponent, PropType } from "vue";
import { Label } from "@/types/Label";

export default defineComponent({
    props: {
        label: { type: Object as PropType<Label>, required: true },
    },
});
</script>

<style scoped>
    .label-container {
        width: 270px;
        height: 150px;
        display: flex;
        flex-direction: column;
        align-items: center;
        border: 1px solid black;
        font-family: Vin Slab Pro Light;
    }
</style>

