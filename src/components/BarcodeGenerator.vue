<template>
    <div>
        <input
            class="button no-print"
            type="button"
            @click="onClick"
            value="Загрузить"
        />
        <span class="no-print">{{
            ` Количество бирок: ${labels.length}`
        }}</span>
        <div v-for="page in pages" :key="page" class="container pagebreak">
            <LabelComponent
                v-for="label in labels.slice(
                    pageSize * page,
                    pageSize * (page + 1)
                )"
                :key="label"
                :label="label"
            />
            <EmptyContainer
                v-for="label in labels.slice(
                    pageSize * page,
                    pageSize * (page + 1)
                )"
                :key="label"
                :label="label"
            />
        </div>
    </div>
</template>

<script lang="ts">
import { computed, defineComponent, ref } from "vue";
import {
    getArrOfExcelFirstWorksheet,
    getExcelFromfile,
    getFileFromLocal,
} from "@/utils/excel";
import { Label } from "@/types/Label";
import LabelComponent from "@/components/LabelComponent.vue";
import EmptyContainer from "./EmptyContainer.vue";

export default defineComponent({
    components: {
        LabelComponent: LabelComponent,
        EmptyContainer: EmptyContainer,
    },
    setup() {
        const pageSize = 30;
        const pages = computed(() => {
            const result: number[] = [];
            for (let i = 0; i < Math.ceil(labels.value.length / pageSize); i++)
                result.push(i);

            return result;
        });

        const labels = ref<Label[]>([]);

        const onClick = async () => {
            labels.value = [];
            await importExcel();
        };

        const importExcel = async () => {
            const file = await getFileFromLocal(".xls, .xlsx", false);

            if (!file) return;

            const wb = await getExcelFromfile(file);
            const data = (await getArrOfExcelFirstWorksheet(wb, {
                header: 1,
            })) as string[][];

            data.forEach((x, i) => {
                if (i !== 0 && x !== []) {
                    if (x[0] !== undefined)
                        labels.value.push({
                            barcode: x[0],
                            description: x[1],
                            size:
                                Number(x[2]).toString() === "NaN"
                                    ? null
                                    : Number(x[2]),
                            footer: x[3],
                        });
                }
            });
        };

        return { onClick, labels, pages, pageSize };
    },
});
</script>

<style scoped>
    .container {
        display: grid;
        grid-template-columns: auto auto auto auto auto;
        justify-content: start;
    }

    .button {
        margin-bottom: 8px;
    }

    @media print {
        .no-print,
        .no-print * {
            display: none !important;
        }
        .pagebreak {
            clear: both;
            page-break-after: always;
        }
        @page {
            size: 420mm 594mm;
        }
    }
</style>
