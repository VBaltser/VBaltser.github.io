<template>
    <div class="label-container">
        <barcode
            class="barcode"
            :value="label.barcode"
            format="EAN13"
            :options="{ width: 3, height: 70, displayValue: true }"
        />
        <span class="description">{{ label.description.toUpperCase() }}</span>
        <span class="size">{{
            label.size === undefined || label.size === null
                ? " "
                : `РАЗМЕР: ${label.size}`
        }}</span>
        <span class="footer">{{ label.footer }}</span>
    </div>
</template>

<script lang="ts">
import { defineComponent, PropType } from "vue";
import VueBarcode from "@chenfengyuan/vue-barcode";
import { Label } from "@/types/Label";

export default defineComponent({
    props: {
        label: { type: Object as PropType<Label>, required: true },
    },
    components: {
        barcode: VueBarcode,
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

.barcode {
    margin-top: 2px;
    width: 90%;
    height: 30%;
}

.footer {
    text-align: center;
    font-size: 12px;
    margin-bottom: 4px;
}

.description {
    margin-top: 12px;
    height: 40%;
    text-align: center;
    font-size: 14px;
    width: 80%;
}

.size {
    font-size: 14px;
    height: 20%;
    text-align: center;
}
</style>

