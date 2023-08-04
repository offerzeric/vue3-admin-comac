export interface ListItem {
  avatar?: string
  title: string
  datetime?: string
  description?: string
  status?: "" | "success" | "info" | "warning" | "danger"
  extra?: string
}

export const notifyData: ListItem[] = [
  {
    title: "COMAC导航数据工程系统已上线，欢迎大家使用！",
    datetime: "1天前",
    status: "info",
    description: "各位同事，COMAC导航数据工程系统已上线，欢迎大家使用！"
  }
  // {
  //   avatar: "https://gw.alipayobjects.com/zos/rmsportal/OKJXDXrmkNshAMvwtvhu.png",
  //   title: "V3 Admin 上线啦",
  //   datetime: "一年前",
  //   description: "一个中后台管理系统基础解决方案，基于 Vue3、TypeScript、Element Plus 和 Pinia"
  // }
]

export const messageData: ListItem[] = [
  // {
  //   avatar: "https://gw.alipayobjects.com/zos/rmsportal/ThXAXghbEsBCCSDihZxY.png",
  //   title: "来自楚门的世界",
  //   description: "如果再也不能见到你，祝你早安、午安和晚安",
  //   datetime: "1998-06-05"
  // }
]

export const todoData: ListItem[] = [
  // {
  //   title: "任务名称",
  //   description: "这家伙很懒，什么都没留下",
  //   extra: "未开始",
  //   status: "info"
  // }
]
