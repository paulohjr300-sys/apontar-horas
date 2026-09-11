/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react-hooks/static-components */
/* eslint-disable react-hooks/immutability */
/* eslint-disable react-hooks/purity */
/* eslint-disable @typescript-eslint/no-explicit-any */
"use client";

import React, { useState, useEffect } from "react";
import { createClient, User } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  Plus,
  Clock,
  Settings,
  LogOut,
  Lock,
  Mail,
  Save,
  Upload,
  BarChart3,
  Pencil,
  Trash2,
  X,
  CheckCircle2,
  XCircle,
  Info,
  Filter,
  CloudCog,
  ShieldCheck,
  CalendarDays,
  CheckCircle,
  ChevronDown,
  ChevronUp,
  Hash,
  Tag,
  FileText,
  Link as LinkIcon,
  CheckSquare,
} from "lucide-react";

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || "";
const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || "";
const supabase = createClient(supabaseUrl, supabaseKey);

interface TimeEntry {
  id: string;
  date: string;
  card_type: string;
  activity: string;
  start_time: string;
  end_time: string;
  card_id: string;
  description: string;
  card_link: string;
  status: "Pendente" | "Lançado";
  ado_doc_id?: string;
}

interface UserSettings {
  work_start_time: string;
  work_end_time: string;
  friday_work_end_time: string;
  lunch_break_minutes: number;
  ado_organization?: string;
  ado_project?: string;
  ado_pat?: string;
  ado_user_name?: string;
  ado_user_id?: string;
}

type ToastType = {
  message: string;
  type: "success" | "error" | "info";
  id: number;
};
type ConfirmDialogType = {
  title: string;
  message: string;
  onConfirm: () => void;
} | null;

export default function Page() {
  const [user, setUser] = useState<User | null>(null);
  const [loadingAuth, setLoadingAuth] = useState(true);

  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authMode, setAuthMode] = useState<"login" | "signup">("login");
  const [isSubmittingAuth, setIsSubmittingAuth] = useState(false);

  const [entries, setEntries] = useState<TimeEntry[]>([]);
  const [loadingData, setLoadingData] = useState(true);
  const [showSettings, setShowSettings] = useState(false);
  const [uploadingExcel, setUploadingExcel] = useState(false);
  const [editingEntry, setEditingEntry] = useState<TimeEntry | null>(null);
  const [isSyncingAdo, setIsSyncingAdo] = useState(false);

  const [collapsedDays, setCollapsedDays] = useState<Record<string, boolean>>(
    {},
  );

  const [selectedMonth, setSelectedMonth] = useState(() =>
    new Date().toISOString().substring(0, 7),
  );

  const [startDateFilter, setStartDateFilter] = useState("");
  const [endDateFilter, setEndDateFilter] = useState("");

  const [toasts, setToasts] = useState<ToastType[]>([]);
  const [confirmDialog, setConfirmDialog] = useState<ConfirmDialogType>(null);

  const [settings, setSettings] = useState<UserSettings>({
    work_start_time: "09:00",
    work_end_time: "19:20",
    friday_work_end_time: "18:20",
    lunch_break_minutes: 60,
    ado_organization: "aguiabranca",
    ado_project: "Lets",
    ado_pat: "",
    ado_user_name: "",
    ado_user_id: "",
  });

  const [form, setForm] = useState({
    date: new Date().toISOString().split("T")[0],
    card_type: "US",
    activity: "Análise e Criação de Cenários de Testes",
    start_time: "09:00",
    end_time: "10:00",
    card_id: "",
    description: "",
    card_link: "",
    status: "Pendente" as "Pendente" | "Lançado",
  });

  const showToast = (
    message: string,
    type: "success" | "error" | "info" = "info",
  ) => {
    const id = Date.now();
    setToasts((prev) => [...prev, { message, type, id }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 4000);
  };

  const addOneHour = (timeStr: string) => {
    if (!timeStr || !timeStr.includes(":")) return "10:00";
    const [h, m] = timeStr.split(":").map(Number);
    const newH = (h + 1) % 24;
    return `${String(newH).padStart(2, "0")}:${String(m || 0).padStart(2, "0")}`;
  };

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setUser(session?.user ?? null);
      setLoadingAuth(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange(
      (_event, session) => {
        setUser(session?.user ?? null);
        setLoadingAuth(false);
      },
    );
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (user) fetchUserData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [user]);

  const fetchUserData = async () => {
    setLoadingData(true);
    try {
      const { data: settingsData } = await supabase
        .from("user_settings")
        .select("*")
        .eq("user_id", user?.id)
        .maybeSingle();

      let loadedStart = "09:00";
      if (settingsData) {
        loadedStart = settingsData.work_start_time || "09:00";
        setSettings({
          work_start_time: loadedStart,
          work_end_time: settingsData.work_end_time || "19:20",
          friday_work_end_time: settingsData.friday_work_end_time || "18:20",
          lunch_break_minutes: settingsData.lunch_break_minutes ?? 60,
          ado_organization: settingsData.ado_organization || "aguiabranca",
          ado_project: settingsData.ado_project || "Lets",
          ado_pat: settingsData.ado_pat || "",
          ado_user_name: settingsData.ado_user_name || "",
          ado_user_id: settingsData.ado_user_id || "",
        });
      }

      const { data: entriesData } = await supabase
        .from("time_entries")
        .select("*")
        .eq("user_id", user?.id)
        .order("date", { ascending: false });

      setEntries(entriesData || []);

      const today = new Date().toISOString().split("T")[0];
      setCollapsedDays({ [today]: true });

      // Preenche dinamicamente o formulário com base no último apontamento do dia de hoje
      if (entriesData && entriesData.length > 0) {
        const todayEntries = entriesData.filter((e) => e.date === today);
        if (todayEntries.length > 0) {
          const lastEntry = todayEntries.sort((a, b) =>
            b.end_time.localeCompare(a.end_time),
          )[0];
          if (lastEntry?.end_time) {
            setForm((prev) => ({
              ...prev,
              date: today,
              start_time: lastEntry.end_time,
              end_time: addOneHour(lastEntry.end_time),
            }));
          }
        }
      }
    } catch (err) {
      showToast("Erro ao carregar os dados.", "error");
    } finally {
      setLoadingData(false);
    }
  };

  const toggleDayCollapse = (dateStr: string) => {
    setCollapsedDays((prev) => ({
      ...prev,
      [dateStr]: !prev[dateStr],
    }));
  };

  const handleAuth = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSubmittingAuth(true);
    if (authMode === "login") {
      const { error } = await supabase.auth.signInWithPassword({
        email,
        password,
      });
      if (error)
        showToast(
          error.message === "Invalid login credentials"
            ? "E-mail ou senha incorretos."
            : error.message,
          "error",
        );
      else showToast("Login realizado com sucesso!", "success");
    } else {
      const { error } = await supabase.auth.signUp({ email, password });
      if (error) showToast(error.message, "error");
      else {
        showToast("Conta criada com sucesso! Você já pode entrar.", "success");
        setAuthMode("login");
      }
    }
    setIsSubmittingAuth(false);
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
    showToast("Sessão encerrada.", "info");
  };

  const handleSaveSettings = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user) return;
    setLoadingData(true);
    const { error } = await supabase.from("user_settings").upsert({
      user_id: user.id,
      work_start_time: settings.work_start_time,
      work_end_time: settings.work_end_time,
      friday_work_end_time: settings.friday_work_end_time,
      lunch_break_minutes: settings.lunch_break_minutes,
      ado_organization: settings.ado_organization,
      ado_project: settings.ado_project,
      ado_pat: settings.ado_pat,
      ado_user_name: settings.ado_user_name,
      ado_user_id: settings.ado_user_id,
      updated_at: new Date().toISOString(),
    });
    if (!error) {
      setShowSettings(false);
      showToast("Configurações salvas com sucesso!", "success");
    } else {
      showToast("Erro ao salvar as configurações.", "error");
    }
    setLoadingData(false);
  };

  const syncToAzureDevOps = async (
    cardId: string,
    start: string,
    end: string,
    activity: string,
    notes: string,
    date: string,
  ) => {
    if (
      !settings.ado_organization ||
      !settings.ado_project ||
      !settings.ado_pat ||
      !settings.ado_user_id
    ) {
      showToast(
        "Preencha todas as suas configurações do Azure (Organização, Projeto, Token e User ID).",
        "error",
      );
      return null;
    }

    setIsSyncingAdo(true);
    try {
      const { data, error } = await supabase.functions.invoke("sync-ado", {
        body: {
          cardId,
          start,
          end,
          activity,
          notes,
          date,
          organization: settings.ado_organization,
          project: settings.ado_project,
          pat: settings.ado_pat,
          userName: settings.ado_user_name,
          userId: settings.ado_user_id,
        },
      });

      if (error) throw new Error(error.message || "Falha na Edge Function");
      if (data?.error) throw new Error(data.error);

      return data?.docId || null;
    } catch (err: any) {
      showToast(err.message || "Erro ao sincronizar com o Azure", "error");
      return null;
    } finally {
      setIsSyncingAdo(false);
    }
  };

  const deleteFromAzureDevOps = async (docId?: string) => {
    if (!docId || !settings.ado_organization || !settings.ado_pat) return;
    try {
      await supabase.functions.invoke("sync-ado", {
        body: {
          action: "DELETE",
          docId,
          organization: settings.ado_organization,
          pat: settings.ado_pat,
        },
      });
    } catch (err) {
      console.error("Erro ao deletar no Azure:", err);
    }
  };

  const parseExcelTime = (val: any): string => {
    if (!val) return "";
    if (val instanceof Date)
      return `${String(val.getHours()).padStart(2, "0")}:${String(val.getMinutes()).padStart(2, "0")}`;
    const valStr = String(val).trim();
    if (valStr.includes(":")) {
      const match = valStr.match(/(\d{1,2}):(\d{2})/);
      if (match)
        return `${String(match[1]).padStart(2, "0")}:${String(match[2]).padStart(2, "0")}`;
    }
    if (typeof val === "number") {
      const totalSeconds = Math.round(val * 86400);
      return `${String(Math.floor(totalSeconds / 3600) % 24).padStart(2, "0")}:${String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, "0")}`;
    }
    return "";
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !user) return;
    setUploadingExcel(true);
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary", cellDates: true });
        const sheetName = wb.SheetNames.includes("Lançamento de Horas")
          ? "Lançamento de Horas"
          : wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rawData: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const rowsToInsert: any[] = [];

        rawData.forEach((row) => {
          if (!row || row.length === 0) return;
          const rawDate = row[0];
          if (
            !rawDate ||
            rawDate === "Data" ||
            String(rawDate).toLowerCase().includes("controle")
          )
            return;

          const cardType = row[1],
            rawStartTime = row[2],
            rawEndTime = row[3];
          const cardId = row[6],
            description = row[7],
            cardLink = row[8],
            statusDevOps = row[9];
          const startTime = parseExcelTime(rawStartTime),
            endTime = parseExcelTime(rawEndTime);

          if (startTime && endTime) {
            let formattedDate = "";
            if (rawDate instanceof Date && !isNaN(rawDate.getTime()))
              formattedDate = rawDate.toISOString().split("T")[0];
            else if (
              typeof rawDate === "string" &&
              rawDate.match(/^\d{4}-\d{2}-\d{2}/)
            )
              formattedDate = rawDate.split("T")[0];
            else if (typeof rawDate === "number")
              formattedDate = new Date(
                Math.round((rawDate - 25569) * 86400 * 1000),
              )
                .toISOString()
                .split("T")[0];

            if (formattedDate) {
              const status =
                statusDevOps === "Lançado" ||
                statusDevOps === "Concluído / Lançado"
                  ? "Lançado"
                  : "Pendente";
              rowsToInsert.push({
                user_id: user.id,
                date: formattedDate,
                card_type: String(cardType || "US"),
                activity: "Outros",
                start_time: startTime,
                end_time: endTime,
                card_id: cardId ? String(cardId) : "",
                description: description ? String(description) : "",
                card_link: cardLink ? String(cardLink) : "",
                status: status,
              });
            }
          }
        });

        if (rowsToInsert.length > 0) {
          const { error } = await supabase
            .from("time_entries")
            .insert(rowsToInsert);
          if (error) showToast("Erro ao importar.", "error");
          else {
            showToast(`${rowsToInsert.length} importados!`, "success");
            fetchUserData();
          }
        } else showToast("Planilha vazia ou formato inválido.", "info");
      } catch (err) {
        showToast("Falha ao processar arquivo Excel.", "error");
      } finally {
        setUploadingExcel(false);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const timeToMinutes = (timeStr: string) => {
    if (!timeStr) return 0;
    const [h, m] = timeStr.split(":").map(Number);
    return h * 60 + m;
  };

  const getEntryDurationMinutes = (start: string, end: string) => {
    const startMin = timeToMinutes(start);
    const endMin = timeToMinutes(end);
    return endMin >= startMin ? endMin - startMin : 1440 - startMin + endMin;
  };

  const toggleStatus = async (entry: TimeEntry) => {
    const newStatus = entry.status === "Lançado" ? "Pendente" : "Lançado";

    if (newStatus === "Lançado" && entry.card_id) {
      const docId = await syncToAzureDevOps(
        entry.card_id,
        entry.start_time,
        entry.end_time,
        entry.activity,
        entry.description,
        entry.date,
      );
      if (docId) {
        setEntries((prev) =>
          prev.map((e) =>
            e.id === entry.id
              ? { ...e, status: "Lançado", ado_doc_id: docId }
              : e,
          ),
        );
        await supabase
          .from("time_entries")
          .update({ status: "Lançado", ado_doc_id: docId })
          .eq("id", entry.id);
        showToast("Time Log gravado na TechsBCN com sucesso!", "success");
      }
      return;
    }

    if (newStatus === "Pendente") {
      if (entry.ado_doc_id) {
        await deleteFromAzureDevOps(entry.ado_doc_id);
      }
      setEntries((prev) =>
        prev.map((e) =>
          e.id === entry.id
            ? { ...e, status: "Pendente", ado_doc_id: undefined }
            : e,
        ),
      );
      await supabase
        .from("time_entries")
        .update({ status: "Pendente", ado_doc_id: null })
        .eq("id", entry.id);
      showToast("Marcado como pendente e removido da TechsBCN.", "info");
    }
  };

  const confirmDelete = (entry: TimeEntry) => {
    setConfirmDialog({
      title: "Excluir Lançamento",
      message:
        "Tem certeza? Isso apagará o registro tanto daqui quanto lá na TechsBCN do Azure.",
      onConfirm: async () => {
        setEntries((prev) => prev.filter((e) => e.id !== entry.id));
        setConfirmDialog(null);

        if (entry.ado_doc_id) {
          await deleteFromAzureDevOps(entry.ado_doc_id);
        }

        const { error } = await supabase
          .from("time_entries")
          .delete()
          .eq("id", entry.id);
        if (error) {
          showToast("Erro ao excluir.", "error");
          fetchUserData();
        } else
          showToast("Excluído com sucesso de ambos os sistemas.", "success");
      },
    });
  };

  const handleUpdateEntry = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingEntry) return;
    setLoadingData(true);

    let newDocId = editingEntry.ado_doc_id || null;
    const isLançado = editingEntry.status === "Lançado";

    if (isLançado && editingEntry.card_id) {
      if (editingEntry.ado_doc_id) {
        await deleteFromAzureDevOps(editingEntry.ado_doc_id);
        newDocId = null;
      }
      newDocId = await syncToAzureDevOps(
        editingEntry.card_id,
        editingEntry.start_time,
        editingEntry.end_time,
        editingEntry.activity,
        editingEntry.description,
        editingEntry.date,
      );
    } else if (!isLançado && editingEntry.ado_doc_id) {
      await deleteFromAzureDevOps(editingEntry.ado_doc_id);
      newDocId = null;
    }

    const { error } = await supabase
      .from("time_entries")
      .update({
        date: editingEntry.date,
        card_type: editingEntry.card_type,
        activity: editingEntry.activity,
        start_time: editingEntry.start_time,
        end_time: editingEntry.end_time,
        card_id: editingEntry.card_id,
        description: editingEntry.description,
        card_link: editingEntry.card_link,
        status: editingEntry.status,
        ado_doc_id: newDocId,
      })
      .eq("id", editingEntry.id);

    if (!error) {
      const updatedObj = { ...editingEntry, ado_doc_id: newDocId || undefined };
      setEntries((prev) =>
        prev.map((e) => (e.id === editingEntry.id ? updatedObj : e)),
      );
      setEditingEntry(null);
      showToast("Lançamento atualizado e sincronizado!", "success");
    } else {
      showToast("Erro ao atualizar o registro.", "error");
    }
    setLoadingData(false);
  };

  const handleSubmitEntry = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user) return;
    setLoadingData(true);

    let docId = null;
    const isLançado = form.status === "Lançado";

    if (isLançado && form.card_id) {
      docId = await syncToAzureDevOps(
        form.card_id,
        form.start_time,
        form.end_time,
        form.activity,
        form.description,
        form.date,
      );
    }

    const { data, error } = await supabase
      .from("time_entries")
      .insert([
        {
          ...form,
          user_id: user.id,
          ado_doc_id: docId,
        },
      ])
      .select();

    if (!error && data) {
      const newEntry = data[0];
      const nextStart = newEntry.end_time;
      const nextEnd = addOneHour(nextStart);

      setEntries([newEntry, ...entries]);
      setForm({
        ...form,
        start_time: nextStart,
        end_time: nextEnd,
        description: "",
        card_id: "",
        card_link: "",
        status: "Pendente",
      });
      setCollapsedDays((prev) => ({ ...prev, [form.date]: true }));
      showToast(
        isLançado
          ? "Time Log salvo na TechsBCN e localmente!"
          : "Atividade registrada no banco local!",
        "success",
      );
    } else {
      showToast("Erro ao registrar a atividade.", "error");
    }
    setLoadingData(false);
  };

  const isDateFriday = (dateStr: string) => {
    if (!dateStr) return false;
    const [y, m, d] = dateStr.split("-").map(Number);
    return new Date(y, m - 1, d, 12, 0, 0).getDay() === 5;
  };

  const getBusinessDaysInMonth = (yearMonth: string) => {
    if (!yearMonth) return { regular: 0, fridays: 0 };
    const [y, m] = yearMonth.split("-").map(Number);
    const daysInMonth = new Date(y, m, 0).getDate();
    let regular = 0;
    let fridays = 0;
    for (let i = 1; i <= daysInMonth; i++) {
      const date = new Date(y, m - 1, i, 12, 0, 0);
      const day = date.getDay();
      if (day >= 1 && day <= 4) regular++;
      else if (day === 5) fridays++;
    }
    return { regular, fridays };
  };

  const regularGrossMinutes = getEntryDurationMinutes(
    settings.work_start_time,
    settings.work_end_time,
  );
  const regularNetMinutes = Math.max(
    0,
    regularGrossMinutes - settings.lunch_break_minutes,
  );
  const fridayGrossMinutes = getEntryDurationMinutes(
    settings.work_start_time,
    settings.friday_work_end_time,
  );
  const fridayNetMinutes = Math.max(
    0,
    fridayGrossMinutes - settings.lunch_break_minutes,
  );

  const { regular: regularDaysThisMonth, fridays: fridaysThisMonth } =
    getBusinessDaysInMonth(selectedMonth);
  const calculatedMonthlyGoal =
    regularDaysThisMonth * (regularNetMinutes / 60) +
    fridaysThisMonth * (fridayNetMinutes / 60);

  const monthFilteredEntries = entries.filter((e) =>
    e.date.startsWith(selectedMonth),
  );

  const totalLoggedMinutes = monthFilteredEntries
    .filter((e) => e.status === "Lançado")
    .reduce(
      (acc, curr) =>
        acc + getEntryDurationMinutes(curr.start_time, curr.end_time),
      0,
    );
  const totalLoggedHours = (totalLoggedMinutes / 60).toFixed(2);

  const filteredEntries = entries.filter((e) => {
    if (!e.date.startsWith(selectedMonth)) return false;
    if (startDateFilter && e.date < startDateFilter) return false;
    if (endDateFilter && e.date > endDateFilter) return false;
    return true;
  });

  const todayStr = new Date().toISOString().split("T")[0];
  const todayNetTarget = isDateFriday(todayStr)
    ? fridayNetMinutes
    : regularNetMinutes;
  const todayLoggedMinutes = entries
    .filter((e) => e.date === todayStr && e.status === "Lançado")
    .reduce(
      (acc, curr) =>
        acc + getEntryDurationMinutes(curr.start_time, curr.end_time),
      0,
    );
  const remainingTodayMinutes = Math.max(
    0,
    todayNetTarget - todayLoggedMinutes,
  );
  const remainingTodayFormatted = `${Math.floor(remainingTodayMinutes / 60)}h ${remainingTodayMinutes % 60}m`;

  const dailySummary = filteredEntries.reduce(
    (acc, entry) => {
      const dateKey = entry.date;
      if (!acc[dateKey]) acc[dateKey] = { totalMin: 0, loggedMin: 0, count: 0 };
      const dur = getEntryDurationMinutes(entry.start_time, entry.end_time);
      acc[dateKey].totalMin += dur;
      if (entry.status === "Lançado") acc[dateKey].loggedMin += dur;
      acc[dateKey].count += 1;
      return acc;
    },
    {} as Record<
      string,
      { totalMin: number; loggedMin: number; count: number }
    >,
  );

  const dailySortedDates = Object.keys(dailySummary).sort((a, b) =>
    b.localeCompare(a),
  );

  if (loadingAuth) {
    return (
      <div className="min-h-screen bg-[#0a0f1d] flex flex-col justify-center items-center gap-4">
        <div className="w-8 h-8 border-2 border-indigo-500/30 border-t-indigo-500 rounded-full animate-spin"></div>
        <span className="text-xs font-semibold text-slate-400 uppercase tracking-widest">
          Carregando Workspace...
        </span>
      </div>
    );
  }

  const GlobalOverlays = () => (
    <>
      <div className="fixed top-6 right-6 z-50 flex flex-col gap-3 pointer-events-none">
        {toasts.map((toast) => (
          <div
            key={toast.id}
            className="pointer-events-auto bg-slate-900/95 backdrop-blur-md border border-slate-800/80 rounded-2xl p-4 flex items-center gap-3.5 shadow-2xl animate-in slide-in-from-right-8 fade-in min-w-[300px]"
          >
            {toast.type === "success" && (
              <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />
            )}
            {toast.type === "error" && (
              <XCircle className="w-5 h-5 text-rose-400 shrink-0" />
            )}
            {toast.type === "info" && (
              <Info className="w-5 h-5 text-indigo-400 shrink-0" />
            )}
            <span className="text-sm font-medium text-slate-200">
              {toast.message}
            </span>
          </div>
        ))}
      </div>
      {confirmDialog && (
        <div className="fixed inset-0 bg-[#0a0f1d]/80 backdrop-blur-sm z-[100] flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl p-6 max-w-sm w-full shadow-2xl space-y-4">
            <h3 className="text-lg font-bold text-white tracking-tight">
              {confirmDialog.title}
            </h3>
            <p className="text-sm text-slate-400 leading-relaxed">
              {confirmDialog.message}
            </p>
            <div className="flex justify-end gap-3 pt-2">
              <button
                onClick={() => setConfirmDialog(null)}
                className="px-4 py-2.5 rounded-xl text-sm font-medium text-slate-400 hover:text-white hover:bg-slate-800/80 transition-colors"
              >
                Cancelar
              </button>
              <button
                onClick={confirmDialog.onConfirm}
                className="px-4 py-2.5 rounded-xl text-sm font-semibold text-white bg-rose-600 hover:bg-rose-500 shadow-lg shadow-rose-600/20 transition-all"
              >
                Sim, excluir
              </button>
            </div>
          </div>
        </div>
      )}
      {isSyncingAdo && (
        <div className="fixed bottom-6 right-6 bg-indigo-600 text-white px-5 py-3 rounded-2xl shadow-2xl flex items-center gap-3.5 z-50 animate-in fade-in slide-in-from-bottom-5">
          <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
          <span className="text-sm font-medium">
            Sincronizando com a TechsBCN...
          </span>
        </div>
      )}
    </>
  );

  if (!user) {
    return (
      <div className="min-h-screen bg-[#0a0f1d] flex flex-col items-center justify-center p-4">
        <GlobalOverlays />
        <div className="w-full max-w-[420px] space-y-8">
          <div className="text-center space-y-3">
            <div className="w-14 h-14 bg-gradient-to-tr from-indigo-600 to-indigo-500 rounded-2xl flex items-center justify-center mx-auto shadow-xl shadow-indigo-500/20 mb-2">
              <Clock className="w-7 h-7 text-white" />
            </div>
            <h1 className="text-2xl font-bold text-white tracking-tight">
              DevOps Hours
            </h1>
            <p className="text-sm text-slate-400">
              Gerencie e sincronize seus apontamentos com eficiência
            </p>
          </div>
          <div className="bg-slate-900/90 backdrop-blur-xl border border-slate-800/80 rounded-3xl p-8 shadow-2xl">
            <form onSubmit={handleAuth} className="space-y-5">
              <div>
                <label className="text-xs font-semibold text-slate-400 block mb-2">
                  E-mail profissional
                </label>
                <div className="relative">
                  <Mail className="w-4 h-4 absolute left-3.5 top-3.5 text-slate-500" />
                  <input
                    type="email"
                    required
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    placeholder="dev@empresa.com"
                    className="w-full bg-slate-950/60 border border-slate-800/80 rounded-xl pl-10 pr-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                  />
                </div>
              </div>
              <div>
                <label className="text-xs font-semibold text-slate-400 block mb-2">
                  Senha
                </label>
                <div className="relative">
                  <Lock className="w-4 h-4 absolute left-3.5 top-3.5 text-slate-500" />
                  <input
                    type="password"
                    required
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder="••••••••"
                    className="w-full bg-slate-950/60 border border-slate-800/80 rounded-xl pl-10 pr-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                  />
                </div>
              </div>
              <button
                type="submit"
                disabled={isSubmittingAuth}
                className="w-full bg-indigo-600 hover:bg-indigo-500 text-white font-semibold py-3 rounded-xl text-sm transition-all shadow-lg shadow-indigo-600/25 mt-2 disabled:opacity-50"
              >
                {isSubmittingAuth
                  ? "Processando..."
                  : authMode === "login"
                    ? "Entrar no Workspace"
                    : "Criar Nova Conta"}
              </button>
            </form>
            <div className="text-center pt-6 mt-6 border-t border-slate-800/60">
              <button
                onClick={() =>
                  setAuthMode(authMode === "login" ? "signup" : "login")
                }
                className="text-sm font-medium text-slate-400 hover:text-indigo-400 transition-colors"
              >
                {authMode === "login"
                  ? "Ainda não tem conta? Cadastre-se"
                  : "Já tem conta? Faça login"}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#0a0f1d] text-slate-200 font-sans selection:bg-indigo-500/30 selection:text-indigo-200">
      <GlobalOverlays />
      <div className="max-w-[1600px] mx-auto p-4 sm:p-6 lg:p-8 space-y-8">
        {/* HEADER */}
        <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 pb-6 border-b border-slate-800/60">
          <div className="flex items-center gap-3.5">
            <div className="w-12 h-12 bg-indigo-600/10 border border-indigo-500/20 rounded-2xl flex items-center justify-center shadow-inner">
              <BarChart3 className="w-6 h-6 text-indigo-400" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-white tracking-tight flex items-center gap-2.5">
                DevOps Hours
                <span className="text-[11px] font-medium bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 px-2.5 py-0.5 rounded-full flex items-center gap-1.5">
                  <ShieldCheck className="w-3.5 h-3.5 text-emerald-400" />{" "}
                  Seguro (Individual)
                </span>
              </h1>
              <p className="text-xs text-slate-400 font-medium mt-0.5">
                {user.email}
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <label className="bg-slate-900/90 border border-slate-800 hover:border-slate-700 hover:bg-slate-800/80 text-slate-200 text-sm font-medium px-4 py-2.5 rounded-xl cursor-pointer transition-all flex items-center gap-2 shadow-sm">
              <Upload className="w-4 h-4 text-indigo-400" />
              <span>{uploadingExcel ? "Importando..." : "Importar Excel"}</span>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleFileUpload}
                disabled={uploadingExcel}
                className="hidden"
              />
            </label>

            <div className="bg-slate-900/90 border border-slate-800 px-4 py-2 rounded-xl flex items-center gap-3 shadow-sm">
              <Clock className="w-4 h-4 text-indigo-400" />
              <div className="flex flex-col">
                <span className="text-[10px] uppercase font-bold text-slate-500 tracking-wider leading-none">
                  Hoje
                </span>
                <span className="font-mono font-semibold text-slate-200 text-xs leading-none mt-1">
                  {loadingData ? "--:--" : remainingTodayFormatted} restantes
                </span>
              </div>
            </div>

            <button
              onClick={() => setShowSettings(!showSettings)}
              className={`p-2.5 bg-slate-900/90 border rounded-xl transition-all shadow-sm ${showSettings ? "border-indigo-500 text-indigo-400 bg-indigo-500/10" : "border-slate-800 text-slate-400 hover:text-white hover:border-slate-700"}`}
              title="Configurações"
            >
              <Settings className="w-4 h-4" />
            </button>

            <button
              onClick={handleLogout}
              className="p-2.5 text-slate-400 hover:text-rose-400 bg-slate-900/90 border border-slate-800 hover:border-rose-900/50 hover:bg-rose-500/10 rounded-xl transition-all shadow-sm"
              title="Sair"
            >
              <LogOut className="w-4 h-4" />
            </button>
          </div>
        </header>

        {/* SETTINGS PANEL */}
        {showSettings && (
          <div className="bg-slate-900/90 backdrop-blur-xl border border-slate-800 rounded-3xl p-6 lg:p-8 shadow-2xl animate-in slide-in-from-top-4 fade-in space-y-6">
            <form onSubmit={handleSaveSettings} className="space-y-6">
              <div>
                <div className="flex items-center justify-between mb-4 border-b border-slate-800/80 pb-3">
                  <h2 className="text-sm font-bold text-white flex items-center gap-2">
                    <Clock className="w-4 h-4 text-indigo-400" /> 1. Jornada
                    Padrão de Trabalho
                  </h2>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Início do Expediente
                    </label>
                    <input
                      type="time"
                      value={settings.work_start_time}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          work_start_time: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Fim (Segunda a Quinta)
                    </label>
                    <input
                      type="time"
                      value={settings.work_end_time}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          work_end_time: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Fim (Sexta-feira)
                    </label>
                    <input
                      type="time"
                      value={settings.friday_work_end_time}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          friday_work_end_time: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Pausa Almoço (min)
                    </label>
                    <input
                      type="number"
                      value={settings.lunch_break_minutes}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          lunch_break_minutes: Number(e.target.value),
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                    />
                  </div>
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between mb-4 border-b border-slate-800/80 pb-3">
                  <h2 className="text-sm font-bold text-white flex items-center gap-2">
                    <CloudCog className="w-4 h-4 text-emerald-400" /> 2. Azure
                    DevOps & TechsBCN (Credenciais Pessoais)
                  </h2>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5 bg-slate-950/60 p-5 rounded-2xl border border-slate-800/80">
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Organização
                    </label>
                    <input
                      type="text"
                      value={settings.ado_organization}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          ado_organization: e.target.value,
                        })
                      }
                      className="w-full bg-slate-900 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Projeto
                    </label>
                    <input
                      type="text"
                      value={settings.ado_project}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          ado_project: e.target.value,
                        })
                      }
                      className="w-full bg-slate-900 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Seu Token (PAT)
                    </label>
                    <input
                      type="password"
                      placeholder="Cole seu PAT aqui"
                      value={settings.ado_pat}
                      onChange={(e) =>
                        setSettings({ ...settings, ado_pat: e.target.value })
                      }
                      className="w-full bg-slate-900 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Seu User ID no Azure
                    </label>
                    <input
                      type="text"
                      value={settings.ado_user_id}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          ado_user_id: e.target.value,
                        })
                      }
                      className="w-full bg-slate-900 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 outline-none transition-all font-mono text-xs"
                    />
                  </div>
                  <div className="lg:col-span-2">
                    <label className="text-xs font-medium text-slate-400 block mb-2">
                      Seu Nome no Azure
                    </label>
                    <input
                      type="text"
                      value={settings.ado_user_name}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          ado_user_name: e.target.value,
                        })
                      }
                      className="w-full bg-slate-900 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                </div>
              </div>

              <div className="flex justify-end pt-2">
                <button
                  type="submit"
                  className="bg-indigo-600 hover:bg-indigo-500 text-white text-sm font-semibold px-6 py-3 rounded-xl transition-all shadow-lg shadow-indigo-600/25 flex items-center gap-2"
                >
                  <Save className="w-4 h-4 text-white" /> Salvar Configurações
                </button>
              </div>
            </form>
          </div>
        )}

        {/* OVERVIEW SECTION & MONTH SELECTOR */}
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 pt-2">
          <h2 className="text-base font-bold text-white tracking-tight">
            Indicadores de Desempenho
          </h2>
          <div className="flex items-center gap-2.5 bg-slate-900/90 border border-slate-800 rounded-xl px-3.5 py-2 focus-within:border-indigo-500 transition-all shadow-sm">
            <Filter className="w-4 h-4 text-indigo-400" />
            <span className="text-xs text-slate-400 font-medium">
              Mês Base Metas:
            </span>
            <input
              type="month"
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="bg-transparent text-sm font-medium text-white outline-none [color-scheme:dark]"
            />
          </div>
        </div>

        {/* 3 CARDS SUPERIORES */}
        <section className="grid grid-cols-1 md:grid-cols-3 gap-5">
          <div className="bg-slate-900/80 backdrop-blur-xl border border-slate-800/80 rounded-3xl p-6 shadow-xl relative overflow-hidden group hover:border-slate-700 transition-all">
            <span className="text-xs font-bold text-slate-400 uppercase tracking-wider block mb-2">
              Total Lançado (Mês)
            </span>
            <div className="flex items-baseline gap-2">
              <span className="text-4xl font-black text-white font-mono tracking-tight">
                {totalLoggedHours}
              </span>
              <span className="text-slate-400 text-sm font-medium">horas</span>
            </div>
            <div className="absolute -right-4 -bottom-4 w-24 h-24 bg-indigo-500/5 rounded-full blur-xl group-hover:bg-indigo-500/10 transition-all"></div>
          </div>

          <div className="bg-slate-900/80 backdrop-blur-xl border border-slate-800/80 rounded-3xl p-6 shadow-xl relative overflow-hidden group hover:border-slate-700 transition-all">
            <div className="flex justify-between items-start mb-2">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider block">
                Meta Dinâmica
              </span>
              <span className="text-[10px] font-semibold bg-slate-800 text-indigo-300 px-2.5 py-1 rounded-full border border-slate-700/60">
                {regularDaysThisMonth} Normais / {fridaysThisMonth} Sextas
              </span>
            </div>
            <div className="flex items-baseline gap-2">
              <span className="text-4xl font-black text-white font-mono tracking-tight">
                {calculatedMonthlyGoal.toFixed(2)}
              </span>
              <span className="text-slate-400 text-sm font-medium">horas</span>
            </div>
            <div className="absolute -right-4 -bottom-4 w-24 h-24 bg-indigo-500/5 rounded-full blur-xl group-hover:bg-indigo-500/10 transition-all"></div>
          </div>

          <div className="bg-slate-900/80 backdrop-blur-xl border border-slate-800/80 rounded-3xl p-6 shadow-xl relative overflow-hidden group hover:border-slate-700 transition-all">
            <span className="text-xs font-bold text-slate-400 uppercase tracking-wider block mb-2">
              Saldo a ser cumprido
            </span>
            <div className="flex items-baseline gap-2">
              <span className="text-4xl font-black text-white font-mono tracking-tight">
                {Math.max(
                  0,
                  calculatedMonthlyGoal - parseFloat(totalLoggedHours),
                ).toFixed(2)}
              </span>
              <span className="text-slate-400 text-sm font-medium">horas</span>
            </div>
            <div className="absolute -right-4 -bottom-4 w-24 h-24 bg-emerald-500/5 rounded-full blur-xl group-hover:bg-emerald-500/10 transition-all"></div>
          </div>
        </section>

        {/* TIME LOG FORM */}
        <section className="bg-slate-900/90 backdrop-blur-xl border border-slate-800 rounded-3xl p-6 lg:p-8 shadow-xl relative overflow-hidden">
          <div className="flex items-center justify-between mb-6 pb-4 border-b border-slate-800/60">
            <h2 className="text-sm font-bold text-white flex items-center gap-2">
              <Plus className="w-4 h-4 text-indigo-400" /> Novo Registro de
              Apontamento
            </h2>
            <span className="text-[11px] font-semibold text-emerald-400 flex items-center gap-1.5 bg-emerald-500/10 px-3 py-1 rounded-full border border-emerald-500/20">
              <ShieldCheck className="w-3.5 h-3.5 text-emerald-400" /> TechsBCN
              Integrado
            </span>
          </div>

          <form
            onSubmit={handleSubmitEntry}
            className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 xl:grid-cols-8 gap-4 relative z-10"
          >
            <div className="lg:col-span-1">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-2 mb-2">
                <CalendarDays className="w-3.5 h-3.5 text-indigo-400" /> Data
              </label>
              <input
                type="date"
                value={form.date}
                onChange={(e) => setForm({ ...form, date: e.target.value })}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
              />
            </div>

            <div className="lg:col-span-2 xl:col-span-2">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-2 mb-2">
                <Clock className="w-3.5 h-3.5 text-indigo-400" /> Período
                (Início / Fim)
              </label>
              <div className="flex gap-2">
                <input
                  type="time"
                  value={form.start_time}
                  onChange={(e) =>
                    setForm({ ...form, start_time: e.target.value })
                  }
                  className="w-1/2 bg-slate-950 border border-slate-800 rounded-xl px-2 py-3 text-sm font-mono text-white text-center focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                />
                <input
                  type="time"
                  value={form.end_time}
                  onChange={(e) =>
                    setForm({ ...form, end_time: e.target.value })
                  }
                  className="w-1/2 bg-slate-950 border border-slate-800 rounded-xl px-2 py-3 text-sm font-mono text-white text-center focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                />
              </div>
            </div>

            <div className="lg:col-span-1">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-2 mb-2">
                <Hash className="w-3.5 h-3.5 text-indigo-400" /> DevOps ID
              </label>
              <input
                type="text"
                placeholder="Ex: 14589"
                value={form.card_id}
                onChange={(e) => setForm({ ...form, card_id: e.target.value })}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm font-mono text-white placeholder:text-slate-600 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
              />
            </div>

            <div className="lg:col-span-2 xl:col-span-1">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-2 mb-2">
                <Tag className="w-3.5 h-3.5 text-indigo-400" /> Tipo
              </label>
              <select
                value={form.card_type}
                onChange={(e) =>
                  setForm({ ...form, card_type: e.target.value })
                }
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-3 text-sm text-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all cursor-pointer"
              >
                <option value="US" className="bg-slate-950 text-white">
                  User Story
                </option>
                <option value="BUG" className="bg-slate-950 text-white">
                  Bug
                </option>
                <option value="Daily" className="bg-slate-950 text-white">
                  Daily
                </option>
                <option value="Planning" className="bg-slate-950 text-white">
                  Planning
                </option>
                <option value="QA Plan" className="bg-slate-950 text-white">
                  QA Plan
                </option>
                                <option value="QA Teste" className="bg-slate-950 text-white">
                  QA Teste
                </option>
                                <option value="QA Aprovado" className="bg-slate-950 text-white">
                  QA Aprovado
                </option>
                                <option value="SOS" className="bg-slate-950 text-white">
                  SOS
                </option>
              </select>
            </div>

            <div className="lg:col-span-2 xl:col-span-3">
              <label className="text-xs font-semibold text-amber-400/90 flex items-center gap-2 mb-2">
                <CheckSquare className="w-3.5 h-3.5 text-amber-400" /> Activity
                (Categoria)
              </label>
              <select
                value={form.activity}
                onChange={(e) => setForm({ ...form, activity: e.target.value })}
                className="w-full bg-slate-950 border border-amber-900/40 rounded-xl px-3.5 py-3 text-sm text-white focus:border-amber-500 focus:ring-1 focus:ring-amber-500 outline-none transition-all cursor-pointer"
              >
                <option
                  value="Daily Scrum / Reunião Diaria"
                  className="bg-slate-950 text-white py-2"
                >
                  Daily Scrum / Reunião diária
                </option>
                <option
                  value="Execução de Testes e Regressão"
                  className="bg-slate-950 text-white py-2"
                >
                  Execução de Testes e Regressão
                </option>
                <option
                  value="Análise e Criação de Cenários de Testes"
                  className="bg-slate-950 text-white py-2"
                >
                  Análise e Criação de Cenários de Testes
                </option>
                <option
                  value="Planning / Refinamento Técnico"
                  className="bg-slate-950 text-white py-2"
                >
                  Planning / Refinamento Técnico
                </option>
                <option
                  value="Code Review / Pull Requests"
                  className="bg-slate-950 text-white py-2"
                >
                  Code Review / Pull Requests
                </option>
                <option
                  value="Coding / Implementação"
                  className="bg-slate-950 text-white py-2"
                >
                  Coding / Implementação
                </option>
                <option
                  value="Correção de Bugs (Bug Fixing)"
                  className="bg-slate-950 text-white py-2"
                >
                  Correção de Bugs (Bug Fixing)
                </option>
              </select>
            </div>

            <div className="lg:col-span-2 xl:col-span-3">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-2 mb-2">
                <FileText className="w-3.5 h-3.5 text-indigo-400" /> Notes
                (Descrição)
              </label>
              <input
                type="text"
                placeholder="Detalhes do que foi executado..."
                value={form.description}
                onChange={(e) =>
                  setForm({ ...form, description: e.target.value })
                }
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white placeholder:text-slate-600 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
              />
            </div>

            <div className="lg:col-span-2 xl:col-span-2">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-2 mb-2">
                <LinkIcon className="w-3.5 h-3.5 text-indigo-400" /> Link URL
                (Opcional)
              </label>
              <input
                type="url"
                placeholder="https://dev.azure.com/..."
                value={form.card_link}
                onChange={(e) =>
                  setForm({ ...form, card_link: e.target.value })
                }
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white placeholder:text-slate-600 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
              />
            </div>

            <div className="lg:col-span-2 xl:col-span-1">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-2 mb-2">
                <ShieldCheck className="w-3.5 h-3.5 text-indigo-400" /> Status
              </label>
              <select
                value={form.status}
                onChange={(e) =>
                  setForm({ ...form, status: e.target.value as any })
                }
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-3 text-sm text-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all cursor-pointer"
              >
                <option value="Pendente" className="bg-slate-950 text-white">
                  Pendente
                </option>
                <option value="Lançado" className="bg-slate-950 text-white">
                  Lançado
                </option>
              </select>
            </div>

            <div className="lg:col-span-2 xl:col-span-2 flex items-end">
              <button
                type="submit"
                disabled={isSyncingAdo}
                className="w-full bg-indigo-600 hover:bg-indigo-500 text-white font-semibold rounded-xl text-sm py-3 transition-all shadow-lg shadow-indigo-600/25 disabled:opacity-50"
              >
                Salvar Time Log
              </button>
            </div>
          </form>
        </section>

        {/* DAILY SUMMARY ACCORDION */}
        <section className="space-y-4">
          <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2">
              <CalendarDays className="w-5 h-5 text-indigo-400" />{" "}
              Acompanhamento Diário por Data
            </h2>

            <div className="flex flex-wrap items-center gap-2 bg-slate-900/90 border border-slate-800 rounded-xl px-3.5 py-2 shadow-sm">
              <span className="text-xs font-semibold text-slate-400">
                Período:
              </span>
              <input
                type="date"
                value={startDateFilter}
                onChange={(e) => setStartDateFilter(e.target.value)}
                className="bg-slate-950 text-xs text-white border border-slate-800 rounded-lg px-2 py-1 outline-none [color-scheme:dark]"
              />
              <span className="text-xs text-slate-500">até</span>
              <input
                type="date"
                value={endDateFilter}
                onChange={(e) => setEndDateFilter(e.target.value)}
                className="bg-slate-950 text-xs text-white border border-slate-800 rounded-lg px-2 py-1 outline-none [color-scheme:dark]"
              />
              {(startDateFilter || endDateFilter) && (
                <button
                  onClick={() => {
                    setStartDateFilter("");
                    setEndDateFilter("");
                  }}
                  className="text-xs text-indigo-400 hover:text-white font-medium ml-2 underline"
                >
                  Limpar
                </button>
              )}
            </div>
          </div>

          {dailySortedDates.length === 0 ? (
            <div className="bg-slate-900/60 border border-slate-800 rounded-3xl p-12 text-center text-slate-500 font-medium">
              Nenhum apontamento cadastrado para este período.
            </div>
          ) : (
            <div className="grid grid-cols-1 gap-3.5">
              {dailySortedDates.map((dateStr) => {
                const dayEntries = filteredEntries.filter(
                  (e) => e.date === dateStr,
                );

                const sortedDayEntries = [...dayEntries].sort((a, b) =>
                  b.end_time.localeCompare(a.end_time),
                );

                const isFriday = isDateFriday(dateStr);
                const dayTargetMin = isFriday
                  ? fridayNetMinutes
                  : regularNetMinutes;
                const totalDayMin = dayEntries.reduce(
                  (acc, curr) =>
                    acc +
                    getEntryDurationMinutes(curr.start_time, curr.end_time),
                  0,
                );
                const loggedDayMin = dayEntries
                  .filter((e) => e.status === "Lançado")
                  .reduce(
                    (acc, curr) =>
                      acc +
                      getEntryDurationMinutes(curr.start_time, curr.end_time),
                    0,
                  );
                const progressPercent = Math.min(
                  100,
                  Math.round((totalDayMin / dayTargetMin) * 100),
                );
                const isComplete = totalDayMin >= dayTargetMin;
                const isExpanded = !!collapsedDays[dateStr];

                return (
                  <div
                    key={dateStr}
                    className="bg-slate-900/80 backdrop-blur-xl border border-slate-800/80 rounded-3xl transition-all shadow-xl overflow-hidden hover:border-slate-700/80"
                  >
                    <div
                      onClick={() => toggleDayCollapse(dateStr)}
                      className="p-5 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 cursor-pointer hover:bg-slate-800/30 transition-colors select-none"
                    >
                      <div className="flex items-center gap-4">
                        <div
                          className={`w-12 h-12 rounded-2xl flex items-center justify-center font-mono font-bold text-sm shrink-0 shadow-inner ${isComplete ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" : "bg-indigo-600/10 text-indigo-400 border border-indigo-500/20"}`}
                        >
                          {dateStr.split("-")[2]}
                        </div>
                        <div>
                          <div className="flex items-center gap-2.5">
                            <h3 className="text-white font-semibold text-sm capitalize">
                              {new Date(
                                dateStr + "T12:00:00",
                              ).toLocaleDateString("pt-BR", {
                                weekday: "long",
                                day: "numeric",
                                month: "long",
                                year: "numeric",
                              })}
                            </h3>
                            {isFriday && (
                              <span className="text-[10px] font-semibold bg-indigo-950/80 text-indigo-400 border border-indigo-800/50 px-2.5 py-0.5 rounded-full">
                                Sexta-feira
                              </span>
                            )}
                          </div>
                          <span className="text-xs text-slate-400 font-medium mt-0.5 block">
                            {dayEntries.length}{" "}
                            {dayEntries.length === 1 ? "registro" : "registros"}{" "}
                            • Lançados: {(loggedDayMin / 60).toFixed(2)}h
                          </span>
                        </div>
                      </div>

                      <div className="flex items-center gap-5 w-full sm:w-auto justify-between sm:justify-end">
                        <div className="text-right">
                          <div className="text-sm font-mono font-bold text-white">
                            {(totalDayMin / 60).toFixed(2)}h{" "}
                            <span className="text-slate-500 font-normal text-xs">
                              / {(dayTargetMin / 60).toFixed(2)}h
                            </span>
                          </div>
                          <span className="text-[10px] font-semibold text-slate-400">
                            {progressPercent}% da meta diária
                          </span>
                        </div>
                        <div className="flex items-center gap-3">
                          {isComplete ? (
                            <CheckCircle className="w-5 h-5 text-emerald-400 shrink-0" />
                          ) : (
                            <div className="w-5 h-5 rounded-full border-2 border-slate-700 shrink-0 flex items-center justify-center">
                              <div className="w-1.5 h-1.5 bg-slate-500 rounded-full"></div>
                            </div>
                          )}
                          <div className="p-1.5 rounded-xl bg-slate-800/80 text-slate-400 hover:text-white transition-colors">
                            {isExpanded ? (
                              <ChevronUp className="w-4 h-4 text-indigo-400" />
                            ) : (
                              <ChevronDown className="w-4 h-4 text-slate-400" />
                            )}
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="w-full bg-slate-950 h-1 rounded-none overflow-hidden">
                      <div
                        className={`h-full transition-all duration-500 ${isComplete ? "bg-emerald-500 shadow-lg shadow-emerald-500/50" : "bg-indigo-500"}`}
                        style={{ width: `${progressPercent}%` }}
                      ></div>
                    </div>

                    {isExpanded && (
                      <div className="p-5 bg-slate-950/60 border-t border-slate-800/80 animate-in fade-in duration-200 space-y-3.5">
                        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                          {sortedDayEntries.map((entry) => {
                            const durMin = getEntryDurationMinutes(
                              entry.start_time,
                              entry.end_time,
                            );
                            const durFormatted = `${Math.floor(durMin / 60)}h ${durMin % 60}m`;
                            const isLançado = entry.status === "Lançado";

                            return (
                              <div
                                key={entry.id}
                                className="bg-slate-900/90 border border-slate-800/80 rounded-2xl p-4 flex flex-col justify-between gap-3.5 hover:border-slate-700 transition-all shadow-sm group"
                              >
                                <div className="space-y-2">
                                  <div className="flex justify-between items-start gap-2">
                                    <span className="font-mono text-xs font-semibold text-white">
                                      {entry.card_link ? (
                                        <a
                                          href={entry.card_link}
                                          target="_blank"
                                          rel="noreferrer"
                                          className="text-indigo-400 hover:underline flex items-center gap-1"
                                        >
                                          #{entry.card_id || "Link"}
                                        </a>
                                      ) : entry.card_id ? (
                                        `#${entry.card_id}`
                                      ) : (
                                        "Sem ID"
                                      )}
                                    </span>
                                    <span className="px-2.5 py-0.5 rounded-md text-[10px] font-semibold bg-slate-800 text-slate-300 border border-slate-700/50">
                                      {entry.card_type}
                                    </span>
                                  </div>
                                  <p className="text-xs text-amber-400 font-semibold line-clamp-1">
                                    {entry.activity}
                                  </p>
                                  <p className="text-xs text-slate-300 line-clamp-2 leading-relaxed">
                                    {entry.description || "Sem descrição"}
                                  </p>
                                </div>

                                <div className="flex items-center justify-between pt-3 border-t border-slate-800/80 text-xs">
                                  <div className="flex items-center gap-2.5 font-mono text-slate-300">
                                    <div className="w-6 h-6 rounded-lg bg-indigo-600/20 border border-indigo-500/30 flex items-center justify-center shrink-0">
                                      <Clock className="w-3.5 h-3.5 text-indigo-400" />
                                    </div>
                                    <span>
                                      {entry.start_time} - {entry.end_time}
                                    </span>
                                    <span className="text-white font-bold">
                                      ({durFormatted})
                                    </span>
                                  </div>
                                  <span
                                    className={`px-2.5 py-0.5 rounded-md text-[10px] font-semibold ${isLançado ? "text-emerald-400 bg-emerald-400/10 border border-emerald-500/20" : "text-amber-400 bg-amber-400/10 border border-amber-500/20"}`}
                                  >
                                    {entry.status}
                                  </span>
                                </div>

                                <div className="flex items-center justify-end gap-1.5 pt-2 border-t border-slate-800/60">
                                  <button
                                    onClick={() => toggleStatus(entry)}
                                    disabled={isSyncingAdo}
                                    className={`px-3 py-1.5 rounded-xl text-xs font-semibold transition-all ${isLançado ? "bg-slate-800/80 text-slate-400 hover:text-white" : "bg-indigo-600/20 text-indigo-400 hover:bg-indigo-600 hover:text-white border border-indigo-500/30"}`}
                                  >
                                    {isLançado
                                      ? "Tornar Pendente"
                                      : "Marcar Lançado"}
                                  </button>
                                  <button
                                    onClick={() => setEditingEntry(entry)}
                                    className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-xl transition-colors"
                                    title="Editar"
                                  >
                                    <Pencil className="w-3.5 h-3.5 text-slate-400" />
                                  </button>
                                  <button
                                    onClick={() => confirmDelete(entry)}
                                    className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-slate-800 rounded-xl transition-colors"
                                    title="Excluir"
                                  >
                                    <Trash2 className="w-3.5 h-3.5 text-slate-400" />
                                  </button>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </section>

        {/* EDIT MODAL */}
        {editingEntry && (
          <div className="fixed inset-0 bg-[#0a0f1d]/80 backdrop-blur-sm flex items-center justify-center p-4 z-50 animate-in fade-in">
            <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-2xl w-full p-6 lg:p-8 shadow-2xl relative">
              <button
                onClick={() => setEditingEntry(null)}
                className="absolute top-6 right-6 p-1.5 text-slate-400 hover:text-white bg-slate-800/50 hover:bg-slate-800 rounded-xl transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
              <h3 className="text-lg font-bold text-white mb-6 tracking-tight">
                Editar Apontamento
              </h3>
              <form onSubmit={handleUpdateEntry} className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="text-xs font-semibold text-slate-400 block mb-2">
                      Data
                    </label>
                    <input
                      type="date"
                      value={editingEntry.date}
                      onChange={(e) =>
                        setEditingEntry({
                          ...editingEntry,
                          date: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white outline-none focus:border-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-semibold text-slate-400 block mb-2">
                      Tipo do Card
                    </label>
                    <select
                      value={editingEntry.card_type}
                      onChange={(e) =>
                        setEditingEntry({
                          ...editingEntry,
                          card_type: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white outline-none focus:border-indigo-500"
                    >
                      <option value="US" className="bg-slate-950 text-white">
                        User Story
                      </option>
                      <option value="BUG" className="bg-slate-950 text-white">
                        Bug
                      </option>
                      <option value="Daily" className="bg-slate-950 text-white">
                        Daily
                      </option>
                      <option
                        value="Planning"
                        className="bg-slate-950 text-white"
                      >
                        Planning
                      </option>
                      <option
                        value="QA Plan"
                        className="bg-slate-950 text-white"
                      >
                        QA Plan
                      </option>
                      <option value="QA Teste" className="bg-slate-950 text-white">
                        QA Teste
                      </option>
                                            <option value="QA Aprovado" className="bg-slate-950 text-white">
                        QA Aprovado
                      </option>
                      <option value="SOS" className="bg-slate-950 text-white">
                        SOS
                      </option>
                    </select>
                  </div>
                </div>

                <div>
                  <label className="text-xs font-semibold text-amber-400 block mb-2">
                    Activity (Categoria)
                  </label>
                  <select
                    value={editingEntry.activity}
                    onChange={(e) =>
                      setEditingEntry({
                        ...editingEntry,
                        activity: e.target.value,
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white outline-none focus:border-indigo-500"
                  >
                    <option
                      value="Daily Scrum / Reunião diaria"
                      className="bg-slate-950 text-white"
                    >
                      Daily Scrum / Reunião diária
                    </option>
                    <option
                      value="Execução de Testes e Regressão"
                      className="bg-slate-950 text-white"
                    >
                      Execução de Testes e Regressão
                    </option>
                    <option
                      value="Análise e Criação de Cenários de Testes"
                      className="bg-slate-950 text-white"
                    >
                      Análise e Criação de Cenários de Testes
                    </option>
                    <option
                      value="Planning / Refinamento Técnico"
                      className="bg-slate-950 text-white"
                    >
                      Planning / Refinamento Técnico
                    </option>
                    <option value="Outros" className="bg-slate-950 text-white">
                      Outros
                    </option>
                  </select>
                </div>

                <div className="grid grid-cols-3 gap-4">
                  <div>
                    <label className="text-xs font-semibold text-slate-400 block mb-2">
                      Início
                    </label>
                    <input
                      type="time"
                      value={editingEntry.start_time}
                      onChange={(e) =>
                        setEditingEntry({
                          ...editingEntry,
                          start_time: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm font-mono text-white outline-none focus:border-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-semibold text-slate-400 block mb-2">
                      Fim
                    </label>
                    <input
                      type="time"
                      value={editingEntry.end_time}
                      onChange={(e) =>
                        setEditingEntry({
                          ...editingEntry,
                          end_time: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm font-mono text-white outline-none focus:border-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-semibold text-slate-400 block mb-2">
                      Status
                    </label>
                    <select
                      value={editingEntry.status}
                      onChange={(e) =>
                        setEditingEntry({
                          ...editingEntry,
                          status: e.target.value as any,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white outline-none focus:border-indigo-500"
                    >
                      <option
                        value="Pendente"
                        className="bg-slate-950 text-white"
                      >
                        Pendente
                      </option>
                      <option
                        value="Lançado"
                        className="bg-slate-950 text-white"
                      >
                        Lançado
                      </option>
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div className="sm:col-span-1">
                    <label className="text-xs font-semibold text-slate-400 block mb-2">
                      ID DevOps
                    </label>
                    <input
                      type="text"
                      value={editingEntry.card_id}
                      onChange={(e) =>
                        setEditingEntry({
                          ...editingEntry,
                          card_id: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm font-mono text-white outline-none focus:border-indigo-500"
                    />
                  </div>
                  <div className="sm:col-span-2">
                    <label className="text-xs font-semibold text-slate-400 block mb-2">
                      Link URL
                    </label>
                    <input
                      type="url"
                      value={editingEntry.card_link}
                      onChange={(e) =>
                        setEditingEntry({
                          ...editingEntry,
                          card_link: e.target.value,
                        })
                      }
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm font-mono text-white outline-none focus:border-indigo-500"
                    />
                  </div>
                </div>

                <div>
                  <label className="text-xs font-semibold text-slate-400 block mb-2">
                    Notes (Descrição)
                  </label>
                  <input
                    type="text"
                    value={editingEntry.description}
                    onChange={(e) =>
                      setEditingEntry({
                        ...editingEntry,
                        description: e.target.value,
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-3 text-sm text-white outline-none focus:border-indigo-500"
                  />
                </div>

                <div className="flex justify-end gap-3 pt-5 mt-2 border-t border-slate-800/80">
                  <button
                    type="button"
                    onClick={() => setEditingEntry(null)}
                    className="px-5 py-2.5 rounded-xl text-sm font-medium text-slate-400 hover:text-white transition-colors"
                  >
                    Cancelar
                  </button>
                  <button
                    type="submit"
                    disabled={isSyncingAdo}
                    className="bg-indigo-600 hover:bg-indigo-500 text-white font-semibold px-6 py-2.5 rounded-xl text-sm shadow-lg shadow-indigo-600/25 transition-all disabled:opacity-50"
                  >
                    Salvar Alterações
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
