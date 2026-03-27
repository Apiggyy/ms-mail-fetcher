import axios from 'axios'

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL ?? ''

const http = axios.create({
  baseURL: API_BASE_URL,
  timeout: 15000,
})

http.interceptors.response.use(
  (response) => response,
  (error) => {
    const { response } = error || {}
    if (!response) {
      return Promise.reject(new Error('网络异常，请稍后重试'))
    }

    const { status, data } = response
    let message = `请求失败: ${status}`

    if (typeof data === 'string' && data.trim()) {
      message = data
    } else if (data?.detail) {
      message = data.detail
    } else if (data?.message) {
      message = data.message
    }

    return Promise.reject(new Error(message))
  },
)

function buildUrl(path, query = {}) {
  const base = API_BASE_URL || window.location.origin
  const url = new URL(path, base)
  Object.entries(query).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(key, String(value))
    }
  })
  return url.toString()
}

export async function request(path, options = {}, query = {}) {
  const cleanQuery = {}
  Object.entries(query).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      cleanQuery[key] = value
    }
  })

  const { body, ...restOptions } = options || {}
  const response = await http.request({
    url: path,
    method: 'GET',
    ...restOptions,
    params: cleanQuery,
    data: body,
  })

  return response.data
}

export { API_BASE_URL, buildUrl }
