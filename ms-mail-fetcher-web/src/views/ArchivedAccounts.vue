<script setup>
import { computed, onMounted, ref } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { CopyDocument, Refresh } from '@element-plus/icons-vue'
import { deleteAccount, exportAccountsUrl, getAccounts, refreshAllAccountTokens } from '@/api/accounts'
import { copyText } from '@/utils/clipboard'

const router = useRouter()
const route = useRoute()

const loading = ref(false)
const refreshAllLoading = ref(false)
const rows = ref([])
const total = ref(0)
const page = ref(1)
const pageSize = ref(10)
const search = ref('')
const selectedRows = ref([])

const thresholdDays = 30

function emitAccountsChanged() {
    if (typeof window !== 'undefined') {
        window.dispatchEvent(new Event('accounts-changed'))
    }
}

const handleCopy = async (text) => {
    if (!text) return
    try {
        await copyText(text)
        ElMessage.success('已复制')
    } catch {
        ElMessage.error('复制失败')
    }
}

const pageCount = computed(() => Math.max(Math.ceil(total.value / pageSize.value), 1))

async function fetchData() {
    loading.value = true
    try {
        const data = await getAccounts({
            is_active: false,
            search: search.value,
            page: page.value,
            page_size: pageSize.value,
        })
        rows.value = data.items || []
        total.value = data.total || 0
        emitAccountsChanged()
    } catch (error) {
        ElMessage.error(error.message || '加载失败')
    } finally {
        loading.value = false
    }
}

function toPositiveInt(value, fallback) {
    const n = Number.parseInt(String(value ?? ''), 10)
    return Number.isFinite(n) && n > 0 ? n : fallback
}

function applyStateFromRouteQuery() {
    page.value = toPositiveInt(route.query.page, 1)
    pageSize.value = toPositiveInt(route.query.page_size, 10)
    search.value = String(route.query.search || '')
}

function syncRouteQuery() {
    const query = {
        page: String(page.value),
        page_size: String(pageSize.value),
    }
    if (search.value) query.search = search.value

    router.replace({
        path: '/archived',
        query,
    })
}

function onSearch() {
    page.value = 1
    syncRouteQuery()
    fetchData()
}

function resetFilters() {
    search.value = ''
    page.value = 1
    syncRouteQuery()
    fetchData()
}

function goMail(row, folder) {
    router.push({
        path: `/mail/${row.id}/${folder}`,
        query: {
            email: row.email,
            _from: '/archived',
            _page: String(page.value),
            _page_size: String(pageSize.value),
            _search: search.value,
        },
    })
}

function onPageChange() {
    syncRouteQuery()
    fetchData()
}

function onExportAll() {
    const url = exportAccountsUrl({
        is_active: false,
    })
    window.open(url, '_blank')
}

async function onDelete(row) {
    try {
        await ElMessageBox.confirm(`确认彻底删除 ${row.email} 吗？此操作不可恢复。`, '危险操作', {
            type: 'error',
            confirmButtonText: '确认删除',
            cancelButtonText: '取消',
        })
        await deleteAccount(row.id)
        ElMessage.success('删除成功')
        if (rows.value.length === 1 && page.value > 1) page.value -= 1
        fetchData()
    } catch {
        // canceled
    }
}

async function onRefreshAllTokens() {
    try {
        await ElMessageBox.confirm('确认一键刷新当前筛选范围内的所有账号 Token 吗？', '提示', {
            type: 'warning',
            confirmButtonText: '确认刷新',
            cancelButtonText: '取消',
        })

        refreshAllLoading.value = true
        const result = await refreshAllAccountTokens({
            is_active: false,
            search: search.value,
        })
        ElMessage.success(
            `刷新完成：总数 ${result.total || 0}，成功 ${result.success || 0}，失败 ${result.failed || 0}`,
        )
        await fetchData()
    } catch (error) {
        if (error?.message) {
            ElMessage.error(error.message)
        }
    } finally {
        refreshAllLoading.value = false
    }
}

function onSelectionChange(selection) {
    selectedRows.value = selection
}

async function onBatchDeleteSelected() {
    if (!selectedRows.value.length) {
        ElMessage.warning('请先勾选账号')
        return
    }

    try {
        await ElMessageBox.confirm(
            `确认彻底删除已选中的 ${selectedRows.value.length} 个账号吗？此操作不可恢复。`,
            '危险操作',
            {
                type: 'error',
                confirmButtonText: '确认删除',
                cancelButtonText: '取消',
            },
        )

        const results = await Promise.allSettled(
            selectedRows.value.map((row) => deleteAccount(row.id)),
        )
        const successCount = results.filter((r) => r.status === 'fulfilled').length
        const failedCount = results.length - successCount

        if (failedCount > 0) {
            ElMessage.warning(`批量删除完成：成功 ${successCount}，失败 ${failedCount}`)
        } else {
            ElMessage.success(`批量删除完成：共 ${successCount} 条`)
        }

        if (rows.value.length === successCount && page.value > 1) page.value -= 1
        selectedRows.value = []
        fetchData()
    } catch {
        // canceled
    }
}

onMounted(() => {
    applyStateFromRouteQuery()
    fetchData()
})
</script>

<template>
    <div class="page-container">
        <el-card shadow="never" class="page-card premium-card">
            <template #header>
                <div class="toolbar">
                    <div class="left">
                        <el-input v-model="search" placeholder="搜索 email / remark" clearable style="width: 260px"
                            @keyup.enter="onSearch" />
                        <el-button type="primary" @click="onSearch">查询</el-button>
                        <el-button @click="resetFilters">重置</el-button>
                        <el-button @click="onExportAll">一键导出所有</el-button>
                        <el-button :loading="refreshAllLoading" @click="onRefreshAllTokens">一键刷新Token</el-button>
                        <el-button type="danger" plain :disabled="!selectedRows.length" @click="onBatchDeleteSelected">
                            批量删除已选（{{ selectedRows.length }}）
                        </el-button>
                    </div>
                </div>
            </template>

            <el-table v-loading="loading" :data="rows" border stripe @selection-change="onSelectionChange"
                style="flex: 1; height: 100%; min-height: 0;">
                <el-table-column type="selection" width="48" />
                <el-table-column label="邮箱" min-width="220">
                    <template #default="{ row }">
                        <div class="copy-cell">
                            <span class="truncate-text">{{ row.email }}</span>
                            <el-button link type="primary" :icon="CopyDocument" @click="handleCopy(row.email)" />
                        </div>
                    </template>
                </el-table-column>
                <el-table-column label="密码" min-width="180">
                    <template #default="{ row }">
                        <div class="copy-cell">
                            <span class="truncate-text">{{ row.password }}</span>
                            <el-button link type="primary" :icon="CopyDocument" @click="handleCopy(row.password)" />
                        </div>
                    </template>
                </el-table-column>
                <el-table-column label="Client ID" min-width="220">
                    <template #default="{ row }">
                        <div class="copy-cell">
                            <span class="truncate-text">{{ row.client_id }}</span>
                            <el-button link type="primary" :icon="CopyDocument" @click="handleCopy(row.client_id)" />
                        </div>
                    </template>
                </el-table-column>
                <el-table-column label="Refresh Token" min-width="280">
                    <template #default="{ row }">
                        <div class="copy-cell">
                            <span class="truncate-text">{{ row.refresh_token }}</span>
                            <el-button link type="primary" :icon="CopyDocument"
                                @click="handleCopy(row.refresh_token)" />
                        </div>
                    </template>
                </el-table-column>
                <el-table-column label="距上次刷新天数" width="140" align="center">
                    <template #default="{ row }">
                        <el-tag :type="row.days_since_refresh > thresholdDays ? 'danger' : 'success'" effect="light"
                            round>
                            {{ row.days_since_refresh }} 天
                        </el-tag>
                    </template>
                </el-table-column>
                <el-table-column prop="remark" label="备注" min-width="180" show-overflow-tooltip />
                <el-table-column label="操作" width="250" fixed="right">
                    <template #default="{ row }">
                        <el-space>
                            <el-button type="primary" link @click="goMail(row, 'inbox')">收信</el-button>
                            <el-button type="warning" link @click="goMail(row, 'spam')">垃圾</el-button>
                            <el-button type="danger" link @click="onDelete(row)">彻底删除</el-button>
                        </el-space>
                    </template>
                </el-table-column>
            </el-table>

            <div class="pager">
                <el-pagination v-model:current-page="page" v-model:page-size="pageSize"
                    layout="total, sizes, prev, pager, next, jumper" :total="total" :page-count="pageCount"
                    :page-sizes="[10, 20, 50, 100]" @change="onPageChange" />
            </div>
        </el-card>
    </div>
</template>

<style scoped>
.page-card {
    flex: 1;
    display: flex;
    flex-direction: column;
    min-height: 600px;
    border: none;
}

:deep(.el-card__body) {
    flex: 1;
    display: flex;
    flex-direction: column;
    min-height: 0;
    overflow: hidden;
    padding-bottom: 0px;
}

.toolbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    flex-wrap: wrap;
}

.left {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
}

.warn {
    color: #f56c6c;
    font-weight: 700;
}

.pager {
    margin-top: 16px;
    display: flex;
    justify-content: flex-end;
    margin-bottom: 20px;
}

.copy-cell {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
}

.truncate-text {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.refresh-btn {
    background: linear-gradient(135deg, #67c23a, #85ce61);
    border: none;
    transition: transform 0.2s, box-shadow 0.2s;
}

.refresh-btn:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(103, 194, 58, 0.4);
}
</style>
