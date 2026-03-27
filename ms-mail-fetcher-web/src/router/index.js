import { createRouter, createWebHistory } from 'vue-router'
import ActiveAccounts from '@/views/ActiveAccounts.vue'
import ArchivedAccounts from '@/views/ArchivedAccounts.vue'
import MailViewer from '@/views/MailViewer.vue'

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      redirect: '/active',
    },
    {
      path: '/active',
      name: 'active',
      component: ActiveAccounts,
    },
    {
      path: '/archived',
      name: 'archived',
      component: ArchivedAccounts,
    },
    {
      path: '/mail/:accountId/:folder',
      name: 'mail-viewer',
      component: MailViewer,
    },
  ],
})

export default router
