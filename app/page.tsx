import Link from "next/link";
import AppHeader from "./app-header";

const dashboardLinks = [
  {
    href: "/bill-difference",
    label: "Water Bill Difference",
    description: "Compare water payment files and review bill differences.",
    category: "Reconciliation",
    mark: "WB",
    color: "bg-cyan-50 text-cyan-700 ring-cyan-100",
  },
  {
    href: "/school-bill-difference",
    label: "School Bill Difference",
    description: "Review differences across school payment records.",
    category: "Reconciliation",
    mark: "SB",
    color: "bg-emerald-50 text-emerald-700 ring-emerald-100",
  },
  {
    href: "/student-import",
    label: "Student Import",
    description: "Import student details from a spreadsheet workbook.",
    category: "Student records",
    mark: "SI",
    color: "bg-amber-50 text-amber-700 ring-amber-100",
  },
];

export default function DashboardPage() {
  return (
    <div className="min-h-screen text-slate-900 md:flex">
      <AppHeader />
      <main className="min-w-0 flex-1">
        <section className="mx-auto max-w-6xl px-5 py-8 sm:px-8 sm:py-12">
          <header className="mb-8 sm:mb-10">
            <p className="text-xs font-bold uppercase tracking-[0.16em] text-teal-700">
              Payment reconciliation
            </p>
            <h2 className="mt-3 text-3xl font-bold tracking-tight text-slate-950 sm:text-4xl">
              Choose a workspace
            </h2>
            <p className="mt-2 max-w-2xl text-sm leading-6 text-slate-500 sm:text-base">
              Open a tool to reconcile payments or manage student records.
            </p>
          </header>

          <nav aria-label="Workspaces">
            <ul className="grid list-none gap-4 p-0 sm:grid-cols-2 xl:grid-cols-3">
              {dashboardLinks.map((item) => (
                <li key={item.href}>
                  <Link
                    href={item.href}
                    className="group flex min-h-44 h-full flex-col rounded-xl border border-slate-200 bg-white p-5 shadow-sm shadow-slate-900/[0.03] transition duration-200 hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-lg hover:shadow-slate-900/[0.07] focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-teal-600 sm:p-6"
                  >
                    <span className="flex items-start justify-between gap-4">
                      <span className="min-w-0">
                        <span className="block text-xs font-semibold uppercase tracking-[0.12em] text-slate-400">
                          {item.category}
                        </span>
                        <span className="mt-2 block text-lg font-bold text-slate-900 group-hover:text-teal-800">
                          {item.label}
                        </span>
                      </span>
                      <span
                        aria-hidden="true"
                        className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-xl text-xs font-black tracking-wide ring-1 ring-inset ${item.color}`}
                      >
                        {item.mark}
                      </span>
                    </span>
                    <span className="mt-3 text-sm leading-6 text-slate-500">
                      {item.description}
                    </span>
                    <span className="mt-auto flex items-center justify-between pt-5 text-sm font-semibold text-teal-700">
                      <span>Open workspace</span>
                      <span aria-hidden="true" className="text-lg transition-transform group-hover:translate-x-1">
                        &rarr;
                      </span>
                    </span>
                  </Link>
                </li>
              ))}
            </ul>
          </nav>
        </section>
      </main>
    </div>
  );
}
